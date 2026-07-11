import numpy as np
import pandas as pd
import re
import os, sys
import sqlalchemy
import requests
import xlsxwriter
import telebot
import getpass
import platform
from dateutil import parser
from datetime import datetime, timedelta
from fast_bitrix24 import Bitrix
from more_itertools.recipes import unique
from sqlalchemy import create_engine
from dotenv import dotenv_values
from mysql.connector import Error
import plotly
import plotly.graph_objs as go
import plotly.express as px
from plotly.subplots import make_subplots
import dash
import dash_bootstrap_components as dbc
from dash import dash_table
from dash import dcc
from dash import html
from dash.dependencies import Input, Output
from dotenv import load_dotenv
import os

# Создание коннекторов данных
load_dotenv()
bot = telebot.TeleBot(os.getenv("TELEGRAM_TOKEN"))
bitra = Bitrix(os.getenv("BITRIX_WEBHOOK"))
chat_id = os.getenv("CHAT_ID")
re_1 = r'[^0-9,.;/]'
pd.set_option('display.max_columns', None)
datename = datetime.now().strftime('%d.%m %H.%M.%S')
tittles = bitra.get_all('crm.lead.fields')                                                                                                                                                               # Получение описания полей лидов
users = pd.DataFrame(bitra.get_all('user.get'))                                                                                                                                                          # Получение списка юзеров
status = pd.DataFrame(bitra.get_all('crm.status.list'))                                                                                                                                                  # Получить справочник CRM форм
status_list = list(filter(None, status[(status['ENTITY_ID'] == "SOURCE")]['NAME'].tolist()))                                                                                                             # Получить список источников
RESULT_all = pd.DataFrame(); RESULT_filter = pd.DataFrame(); RESULT_filter_call = pd.DataFrame(); RESULT_filter_message = pd.DataFrame(); RESULT_err = pd.DataFrame()                                    # Создание пустого dataframe
department_list = ['Косметология', 'Пластическая хирургия', 'Рентген', 'Стоматология', 'Флебология']                                                                                                     # Список отделений для декомпозиции

#Создание словарей для заголовков Bitrix24
dict_2 = dict((tit['ID'], tit['VALUE']) for tit in tittles['UF_CRM_1716553574']['items'])                                                                                                                # Словарь для поля 'Источник рекламы'
dict_4 = dict((tit['ID'], tit['VALUE']) for tit in tittles['UF_CRM_1718965625']['items'])                                                                                                                # Словарь для поля 'Причина не записи пациента на консультацию'
dict_5 = dict((tit['ID'], tit['VALUE']) for tit in tittles['UF_CRM_1722587481']['items'])                                                                                                                # Словарь для поля 'Доктор в обращении'
dict_6 = dict((tit['ID'], tit['VALUE']) for tit in tittles['UF_CRM_1722587572']['items'])                                                                                                                # Словарь для поля 'Отделение в обращении'
def Input_lag():
    global start_date, end_date
    while True:
        try:
            start_date = input("Введите дату начала в формате dd.mm.yy:\n")
            start_date = pd.to_datetime(datetime.strptime(start_date, '%d.%m.%y'))
            print("Начальная дата: " + str(start_date)); break
        except: print("Введённый параметр не соответствует допустимому формату даты - попробуйте написать дату по другому")
    while True:
        try:
            end_date = input("Теперь введите конечную дату (обратите внимание, что по умолчанию дате присвоится время 00:00:00, т.е. если вы хотите посчитать, например, до 10.10.24 ВКЛЮЧИТЕЛЬНО, то следует написать 11.10.24):\n")
            if end_date == "": end_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0); print(end_date)
            else: end_date = pd.to_datetime(datetime.strptime(end_date, '%d.%m.%y'))
            if end_date < start_date: print("Конечная дата меньше начальной, так быть не должно, попробуйте снова"); continue
            print("Конечная дата: " + str(end_date)); break
        except: print("Введённый параметр не соответствует допустимому формату даты - попробуйте написать дату по другому")
    print("Все опциональные параментры введены пользователем. Ожидаем дозагрузки данных и начала расчёта...")
def SendTelegram(status, er):
	'''
    UserName = getpass.getuser()                                                                                                                                                                         # Имя пользователя (обычно оно User - не информативно)
    CompName = platform.node()                                                                                                                                                                           # Имя компьютера
    chat_id = '5249664773'                                                                                                                                                                               # ID моей телеги
    if status == "try": # Если связь с телегой установлена
        bot.send_message(chat_id, datename+" пользователь "+UserName+" ("+CompName+") успешно воспользовался скриптом оценки конверсии КЦ")                                                                  # Отправка сообщения
        photo = open("Конверсия КЦ " + datename + "/Инфографика/Конверсия лидов по отделам.png", "rb")
        bot.send_photo(chat_id, photo)
    elif status == "except1": # Если нет подключения к SQL серверу
        bot.send_message(chat_id, "ERROR: " + datename+" пользователь "+UserName+" ("+CompName+") неудачно запустил скрипт оценки конверсии КЦ: " + er)                                                      # Отправка сообщения
    '''
def DataConcat(lists, datefirst, datelast):
    DataLeads = pd.DataFrame(bitra.get_all('crm.lead.list', params={'select': ['ID'], 'filter': {'>DATE_CREATE': str(datefirst), '<DATE_CREATE': str(datelast)}}))                                       # Создание фрейма данных
    for col in list(lists):                                                                                                                                                                              # Вводные для progress bar и запуск цикла по excel файлам
        Colomn = pd.DataFrame(bitra.get_all('crm.lead.list', params={'select': ['ID', col], 'filter': {'>DATE_CREATE': datefirst, '<DATE_CREATE': datelast}}))
        DataLeads = DataLeads.merge(Colomn, on='ID')                                                                                                                                                     # Объединение нескольких excel файлов в один фрейм
    DataLeads = DataLeads.reset_index(drop=True).sort_values(by=['ID'], ascending=False)                                                                                                                 # Сброс индексов и сортировка по ID
    return DataLeads
def Get_Dict(X):                                                                                                                                                                                         # Получение телефона из словаря
    try: return X[0]['VALUE']
    except: pass
def Data_Departament(dep, base, BD):
    BD_C = BD[(BD['Статус'] == 'Перешёл в сделку')]
    BD_group = BD.groupby('Оператор').agg({'ID лида': ['count']}).reset_index(); BD_group.columns = ['Оператор', 'Количество всех лидов']                                                                # Группируем по операторам из всех лидов
    BD_group_C = BD_C.groupby('Оператор').agg({'ID лида': ['count']}).reset_index(); BD_group_C.columns = ['Оператор', 'Количество лидов, перешедших в сделки']                                          # Группируем по операторам из лидов, перешедших в сделки
    Base_operator = BD_group.merge(BD_group_C, on='Оператор', how='outer')                                                                                                                               # Объединение группированных таблиц
    Base_operator = pd.concat([Base_operator, pd.DataFrame({'Оператор': 'Весь КЦ, отделение: ' + dep, 'Количество всех лидов': Base_operator['Количество всех лидов'].sum(), 'Количество лидов, перешедших в сделки': Base_operator['Количество лидов, перешедших в сделки'].sum()}, index=[0])], axis=0)
    Base_operator['% конверсии'] = round(Base_operator['Количество лидов, перешедших в сделки']/Base_operator['Количество всех лидов'], 3)                                                               # Подсчёт конверсии с округлением round
    Base_operator = Base_operator.fillna(0)                                                                                                                                                              # Чтобы результаты по отделениям были 0%
    Base_operator = Base_operator.reset_index(drop=True)                                                                                                                                                 # Сброс индекса строк (дублируются из-за конкатенации)
    #Base_operator.replace('NaN', np.nan, inplace=True); Base_operator.replace(0, np.nan, inplace=True)                                                                                                  # Замена нулей на NAN
    Base_operator = pd.DataFrame({'Оператор': list(filter(None, base['Оператор'].unique().tolist()))}).merge(Base_operator, on='Оператор', how='outer')                                                  # Объединяем таблицу со списком операторов (для одинаковой структуры)
    Base_operator = Base_operator.fillna(0)
    Base_operator['Отделение'] = dep                                                                                                                                                                     # Добавляет название отделения
    Base_operator = pd.concat([Base_operator, pd.DataFrame({'Оператор': np.nan, 'Количество всех лидов': np.nan, 'Количество лидов, перешедших в сделки': np.nan, '% конверсии': np.nan}, index=[0])], axis=0) # Добавление промежуточной строки между отделениями

    return Base_operator
def Convertions(base):
    department = pd.DataFrame()
    for dep in department_list:
        collect = Data_Departament(dep, base, base.loc[(base['Отделение'] == dep)])
        department = pd.concat([department, collect], axis=0)
    collect = Data_Departament("Все отделения", base, base)
    department = pd.concat([department, collect], axis=0)
    collect = department.dropna()[(department.dropna()['Оператор'].str.contains('Весь КЦ'))]
    department = pd.concat([department, collect], axis=0)
    return department
def Data_source(dep, f, BD):
    BD_C = BD[(BD['Статус'] == 'Перешёл в сделку')]
    BD_group_C = BD_C.groupby('Источник лида').agg({'ID лида': ['count']}).reset_index(); BD_group_C.columns = ['Источник лида', str(f)]                                                                 # Группируем по операторам из лидов, перешедших в сделки
    Base_operator = pd.DataFrame({'Источник лида': status_list}).merge(BD_group_C, on='Источник лида', how='outer')                                                                                      # Объединяем таблицу со списком операторов (для одинаковой структуры)
    Base_operator = Base_operator.fillna(0)
    print(Base_operator)
    return Base_operator
def Convertions_source(base):
    global start_date, end_date
    source = pd.DataFrame()
    print(start_date)
    print(end_date)
    first_date_list = pd.date_range(start_date.replace(hour=0, minute=0, second=0), end_date.replace(hour=0, minute=0, second=0), freq="D").tolist()
    end_date_list = pd.date_range(start_date.replace(hour=0, minute=0, second=0), end_date.replace(hour=0, minute=0, second=0), freq="D").tolist()
    firstin = pd.to_datetime(first_date_list[:-1]); endin = pd.to_datetime(end_date_list[1:])
    print(firstin)
    print(endin)
    for dep in department_list:
        for f, e in zip(firstin, endin):
            f = pd.to_datetime(f); e = pd.to_datetime(e)
            print(type(f)); print(type(e))
            #base['Дата создания'] = pd.to_datetime(base['Дата создания'], format='%Y-%m-%d %H:%M:%S')
            #base['Дата создания'] = base['Дата создания'].apply(datetime.strptime, args=('%d-%b-%Y %H:%M:%S',))
            #base['Дата создания'] = pd.Timestamp(base['Дата создания'])
            #base['Дата создания'] = base['Дата создания'].strftime('%Y-%m-%d %X')
            print(base['Дата создания'])
            print(type(base['Дата создания'][1]))
            collect = Data_source(dep, f, base.loc[((base['Отделение'] == dep) & (base['Дата создания'] >= f) & (base['Дата создания'] < e))])
            #source = source.merge(collect, on='Оператор', how='outer')
            source = pd.concat([source, source], axis=1)
            print(source)

    Base_operator = BD_group.merge(BD_group_C, on='Оператор', how='outer')


    collect = Data_source("Все отделения", base, base)
    department = pd.concat([department, collect], axis=0)
    collect = department.dropna()[(department.dropna()['Оператор'].str.contains('Весь КЦ'))]
    department = pd.concat([department, collect], axis=0)
    return department
def Make_plot():
    ###################################################################################################################################################################################################
    # График Конверсия звонков по отделам
    collect = Conv_call.drop_duplicates()
    collect = collect.dropna()[(collect.dropna()['Оператор'].str.contains('Весь КЦ'))]
    fig = go.Figure({"layout": {"title": {"text": "Конверсия звонков по отделам"}}})
    fig.add_trace(go.Bar(x=collect['Оператор'], y=collect['Количество лидов, перешедших в сделки'], text=(collect['% конверсии']*100).round(2).mask(((collect['% конверсии']*100).round(2)) == 0, ""), texttemplate = "%{text}%", textposition='outside', textfont=dict(weight="bold", size=10), name='Сконвертированы в сделки', marker_color='green'))
    fig.add_trace(go.Bar(x=collect['Оператор'], y=(collect['Количество всех лидов'] - collect['Количество лидов, перешедших в сделки']), text=((100-collect['% конверсии']*100).round(2)).mask(((100-collect['% конверсии']*100).round(2)) == 100, ""), texttemplate="%{text}%", textposition='outside', textfont=dict(weight="bold", size=10), name='Не сконвертированы'))
    fig.update_layout(barmode='stack', legend=dict(yanchor="top", y=0.99, xanchor="left", x=0.01)) #, legend_orientation="center"
    fig.write_image("Конверсия КЦ " + datename + "/Инфографика/Конверсия звонков по отделам.png", format="png", width=600, height=600, scale=1)
    # График Конверсия звонков по специалистам
    collect = Conv_call.dropna()[((Conv_call.dropna()['Отделение'] == 'Все отделения')) & (~Conv_call.dropna()['Оператор'].str.contains('Весь КЦ'))]
    fig = go.Figure({"layout": {"title": {"text": "Конверсия звонков по операторам"}}})
    fig.add_trace(go.Bar(x=collect['Оператор'], y=collect['Количество лидов, перешедших в сделки'], text=(collect['% конверсии'] * 100).round(2).mask(((collect['% конверсии']*100).round(2)) == 0, ""), texttemplate="%{text}%", textposition='outside', textfont=dict(weight="bold", size=10), name='Сконвертированы в сделки', marker_color='green'))
    fig.add_trace(go.Bar(x=collect['Оператор'], y=(collect['Количество всех лидов'] - collect['Количество лидов, перешедших в сделки']), text=((100-collect['% конверсии']*100).round(2)).mask(((100-collect['% конверсии']*100).round(2)) == 100, ""), texttemplate="%{text}%", textposition='outside', textfont=dict(weight="bold", size=10), name='Не сконвертированы'))
    fig.update_layout(barmode='stack', legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
    fig.write_image("Конверсия КЦ " + datename + "/Инфографика/Конверсия звонков по операторам.png", format="png", width=600, height=600, scale=1)

    ###################################################################################################################################################################################################
    # График Конверсия сообщений по отделам
    collect = Conv_message.drop_duplicates()
    collect = collect.dropna()[(collect.dropna()['Оператор'].str.contains('Весь КЦ'))]
    fig = go.Figure({"layout": {"title": {"text": "Конверсия сообщений по отделам"}}})
    fig.add_trace(go.Bar(x=collect['Оператор'], y=collect['Количество лидов, перешедших в сделки'], text=(collect['% конверсии']*100).round(2).mask(((collect['% конверсии']*100).round(2)) == 0, ""), texttemplate = "%{text}%", textposition='outside', textfont=dict(weight="bold", size=10), name='Сконвертированы в сделки', marker_color='green'))
    fig.add_trace(go.Bar(x=collect['Оператор'], y=(collect['Количество всех лидов'] - collect['Количество лидов, перешедших в сделки']), text=((100-collect['% конверсии']*100).round(2)).mask(((100-collect['% конверсии']*100).round(2)) == 100, ""), texttemplate="%{text}%", textposition='outside', textfont=dict(weight="bold", size=10), name='Не сконвертированы'))
    fig.update_layout(barmode='stack', legend=dict(yanchor="top", y=0.99, xanchor="left", x=0.01)) #, legend_orientation="center"
    fig.write_image("Конверсия КЦ " + datename + "/Инфографика/Конверсия сообщений по отделам.png", format="png", width=600, height=600, scale=1)
    # График Конверсия сообщений по специалистам
    collect = Conv_message.dropna()[((Conv_message.dropna()['Отделение'] == 'Все отделения')) & (~Conv_message.dropna()['Оператор'].str.contains('Весь КЦ'))]
    fig = go.Figure({"layout": {"title": {"text": "Конверсия сообщений по операторам"}}})
    fig.add_trace(go.Bar(x=collect['Оператор'], y=collect['Количество лидов, перешедших в сделки'], text=(collect['% конверсии'] * 100).round(2).mask(((collect['% конверсии']*100).round(2)) == 0, ""), texttemplate="%{text}%", textposition='outside', textfont=dict(weight="bold", size=10), name='Сконвертированы в сделки', marker_color='green'))
    fig.add_trace(go.Bar(x=collect['Оператор'], y=(collect['Количество всех лидов'] - collect['Количество лидов, перешедших в сделки']), text=((100-collect['% конверсии']*100).round(2)).mask(((100-collect['% конверсии']*100).round(2)) == 100, ""), texttemplate="%{text}%", textposition='outside', textfont=dict(weight="bold", size=10), name='Не сконвертированы'))
    fig.update_layout(barmode='stack', legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
    fig.write_image("Конверсия КЦ " + datename + "/Инфографика/Конверсия сообщений по операторам.png", format="png", width=600, height=600, scale=1)

    ###################################################################################################################################################################################################
    # График Конверсия лидов по отделам
    collect = Conv_all.drop_duplicates()
    collect = collect.dropna()[(collect.dropna()['Оператор'].str.contains('Весь КЦ'))]
    fig = go.Figure({"layout": {"title": {"text": "Конверсия лидов по отделам"}}})
    fig.add_trace(go.Bar(x=collect['Оператор'], y=collect['Количество лидов, перешедших в сделки'], text=(collect['% конверсии']*100).round(2).mask(((collect['% конверсии']*100).round(2)) == 0, ""), texttemplate = "%{text}%", textposition='outside', textfont=dict(weight="bold", size=10), name='Сконвертированы в сделки', marker_color='green'))
    fig.add_trace(go.Bar(x=collect['Оператор'], y=(collect['Количество всех лидов'] - collect['Количество лидов, перешедших в сделки']), text=((100-collect['% конверсии']*100).round(2)).mask(((100-collect['% конверсии']*100).round(2)) == 100, ""), texttemplate="%{text}%", textposition='outside', textfont=dict(weight="bold", size=10), name='Не сконвертированы'))
    fig.update_layout(barmode='stack', legend=dict(yanchor="top", y=0.99, xanchor="left", x=0.01)) #, legend_orientation="center"
    fig.write_image("Конверсия КЦ " + datename + "/Инфографика/Конверсия лидов по отделам.png", format="png", width=600, height=600, scale=1)
    # График Конверсия лидов по специалистам
    collect = Conv_all.dropna()[((Conv_all.dropna()['Отделение'] == 'Все отделения')) & (~Conv_all.dropna()['Оператор'].str.contains('Весь КЦ'))]
    fig = go.Figure({"layout": {"title": {"text": "Конверсия лидов по операторам"}}})
    fig.add_trace(go.Bar(x=collect['Оператор'], y=collect['Количество лидов, перешедших в сделки'], text=(collect['% конверсии'] * 100).round(2).mask(((collect['% конверсии']*100).round(2)) == 0, ""), texttemplate="%{text}%", textposition='outside', textfont=dict(weight="bold", size=10), name='Сконвертированы в сделки', marker_color='green'))
    fig.add_trace(go.Bar(x=collect['Оператор'], y=(collect['Количество всех лидов'] - collect['Количество лидов, перешедших в сделки']), text=((100-collect['% конверсии']*100).round(2)).mask(((100-collect['% конверсии']*100).round(2)) == 100, ""), texttemplate="%{text}%", textposition='outside', textfont=dict(weight="bold", size=10), name='Не сконвертированы'))
    fig.update_layout(barmode='stack', legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
    fig.write_image("Конверсия КЦ " + datename + "/Инфографика/Конверсия лидов по операторам.png", format="png", width=600, height=600, scale=1)

    ###################################################################################################################################################################################################
    # Ошибки (по операторам)
    collect = RESULT_err.groupby('Оператор').agg({'Описание ошибки': ['count']}).reset_index(); collect.columns = ['Оператор', 'Количество логических ошибок']
    fig = go.Figure({"layout": {"title": {"text": "Распределение логических ошибок по операторам"}}})
    fig.add_trace(go.Bar(x=collect['Оператор'], y=collect['Количество логических ошибок'], text=collect['Количество логических ошибок'], textposition='outside', textfont=dict(weight="bold", size=10), name='Ошибки'))
    fig.write_image("Конверсия КЦ " + datename + "/Инфографика/Распределение логических ошибок по операторам.png", format="png", width=600, height=600, scale=1)
    # Ошибки (по причинам)

    collect = RESULT_err.groupby('Описание ошибки').agg({'Оператор': ['count']}).reset_index(); collect.columns = ['Описание ошибки', 'Количество логических ошибок']
    pull = [0] * len(collect['Описание ошибки'])
    try: pull[collect['Описание ошибки'].tolist().index(collect['Описание ошибки'].max())] = 0.2
    except: pull = 0
    fig = go.Figure({"layout": {"title": {"text": "Распределение логических ошибок по причинам"}}})
    fig.add_trace(go.Pie(values=collect['Количество логических ошибок'], text=collect['Количество логических ошибок'], labels=collect['Описание ошибки'], pull=pull, hole=0.5, textfont=dict(weight="bold", size=10)))
    fig.update_layout(annotations=[dict(text='Всего<br>ошибок:<br>'+str(RESULT_err['Описание ошибки'].count()), x=0.5, y=0.5, font_size=14, showarrow=False)], legend=dict(orientation="h"))
    fig.write_image("Конверсия КЦ " + datename + "/Инфографика/Распределение логических ошибок по причинам.png", format="png", width=600, height=600, scale=1)
def ContactID(ID):
    try:
        contact = bitra.get_by_ID('crm.contact.get', [ID])
        if not contact:
            return pd.Series([None, None, None])
        if isinstance(contact, dict):
            if "result" in contact: data = contact["result"]
            else: data = contact.get(str(ID), contact)
        else: data = contact
        phones = []
        if isinstance(data, dict) and "PHONE" in data: phones = [p.get("VALUE") for p in data["PHONE"] if isinstance(p, dict) and "VALUE" in p]
        phones = phones + [None] * (3 - len(phones))
        if len(phones) > 3: phones[2] = ", ".join([p for p in phones[2:] if p])
        return pd.Series(phones[:3])
    except Exception as e:
        print(f"Ошибка ContactID для {ID}: {e}")
        return pd.Series([None, None, None])
def clean_phone(val):
    if isinstance(val, list):
        nums = [v.get("VALUE") for v in val if isinstance(v, dict) and "VALUE" in v]
        val = nums[0] if nums else None
    elif isinstance(val, dict): val = val.get("VALUE")
    return re.sub(re_1, '', str(val)) if val else None
def BitrixFrame(datefirst, datelast):  # Получение данных из Bitrix24
    global RESULT_all, RESULT_filter, RESULT_filter_call, RESULT_filter_message, RESULT_err, DataLeads_Status_NEW
    DataLeads = DataConcat([
        'CONTACT_ID', 'SOURCE_ID', 'STATUS_ID', 'PHONE', 'ASSIGNED_BY_ID', 'DATE_CREATE', 'UF_CRM_1718965625', 'UF_CRM_1716553574', 'UF_CRM_1722587481', 'UF_CRM_1722587572'],
        datefirst, datelast)
    #DataLeads = DataLeads[['CONTACT_ID', 'SOURCE_ID', 'STATUS_ID', 'NAME', 'SECOND_NAME', 'LAST_NAME', 'PHONE', 'ASSIGNED_BY_ID', 'DATE_CREATE', 'UF_CRM_1718965625', 'UF_CRM_1716553574', 'UF_CRM_1722587481', 'UF_CRM_1722587572']]
    #DataLeads.columns = ['CONTACT_ID', 'SOURCE_ID', 'STATUS_ID', 'Имя клиента', 'Фамилия клиента', 'Отчество клиента', 'PHONE', 'ASSIGNED_BY_ID', 'DATE_CREATE', 'UF_CRM_1718965625', 'UF_CRM_1716553574', 'UF_CRM_1722587481', 'UF_CRM_1722587572']
    phones_df = DataLeads['CONTACT_ID'].apply(lambda x: ContactID(x) if x else pd.Series([None, None, None]))
    phones_df.columns = ["Телефон 1", "Телефон 2", "Телефон 3"]
    DataLeads = pd.concat([DataLeads, phones_df], axis=1)
    
    # Применение словарей для столбцов (по умолчанию выгружаются ID показателей словарей)
    DataLeads['UF_CRM_1716553574'] = DataLeads['UF_CRM_1716553574'].replace(dict_2)                                                                                                                      # Присвоение значений для поля 'Источник рекламы'
    DataLeads['UF_CRM_1718965625'] = DataLeads['UF_CRM_1718965625'].replace(dict_4)                                                                                                                      # Присвоение значений для поля 'Причина не записи пациента на консультацию'
    DataLeads['UF_CRM_1722587481'] = DataLeads['UF_CRM_1722587481'].replace(dict_5)                                                                                                                      # Присвоение значений для поля 'Доктор в обращении'
    DataLeads['UF_CRM_1722587572'] = DataLeads['UF_CRM_1722587572'].replace(dict_6)                                                                                                                      # Присвоение значений для поля 'Отделение в обращении'
    DataLeads = DataLeads.merge(users[['ID', 'NAME', 'LAST_NAME', 'WORK_POSITION']], how='left', left_on='ASSIGNED_BY_ID', right_on='ID')                                                                # Присоединение данных о пользователях
    DataLeads = DataLeads.merge(status[['STATUS_ID', 'NAME']], how='left', left_on='SOURCE_ID', right_on='STATUS_ID')                                                                                    # Присоединение данных о источниках
    DataLeads = DataLeads.merge(status[['STATUS_ID', 'NAME']], how='left', left_on='STATUS_ID_x', right_on='STATUS_ID', suffixes=('_x1', '_x2'))                                                         # Присоединение данных о статусах лидов
    DataLeads['Оператор'] = (DataLeads['NAME_x'].astype(str) + " " + DataLeads['LAST_NAME'].astype(str)).str.strip()
    
    DataLeads['PHONE'] = DataLeads['PHONE'].apply(lambda v: (next((d.get('VALUE') for d in v if isinstance(d, dict) and 'VALUE' in d), None) if isinstance(v, list) else (v.get('VALUE') if isinstance(v, dict) else v)))
    for col in ["PHONE", "Телефон 1", "Телефон 2", "Телефон 3"]:
	    if col in DataLeads.columns: DataLeads[col] = DataLeads[col].apply(clean_phone)
    
    
    
    
    
    DataLeads = DataLeads.drop_duplicates(subset=["ID_x"], keep="first")
    
    DataLeads = DataLeads[['ID_x', 'DATE_CREATE', 'Оператор', 'WORK_POSITION', 'NAME', 'UF_CRM_1722587481', 'UF_CRM_1722587572', 'UF_CRM_1718965625', 'UF_CRM_1716553574', 'NAME_y', 'PHONE', 'Телефон 1', 'Телефон 2', 'Телефон 3']]
    DataLeads.columns = ['ID лида', 'Дата создания', 'Оператор', 'Должность', 'Статус', 'На кого обращение', 'Отделение', 'Причина отказа', 'Источник рекламы', 'Источник лида', 'Телефон (если в лиде)', 'Телефон 1 (если в карточке)', 'Телефон 2 (если в карточке)', 'Телефон 3 (если в карточке)']
    DataLeads_Status = DataLeads.copy()
    DataLeads_Status['Все обращения'] = DataLeads_Status['Статус'].mask(DataLeads_Status['Статус'] != 'Перешёл в сделку', 'Не перешёл в сделку')
    # Работа с ошибками
    DataLeads_Status_err_1 = DataLeads_Status[(((DataLeads_Status['Причина отказа'].notnull()) & (DataLeads_Status['Причина отказа'] != "Выбрал другую клинику/доктора") &                               # У сконвертированной сделки пристуствует причина отказа
                                    (DataLeads_Status['Причина отказа'] != "Дорого/ не устраивает уровень цен") & (DataLeads_Status['Причина отказа'] != "Не отвечает на прогревы") &
                                    (DataLeads_Status['Причина отказа'] != "Отказался от дальнейшей связи") & (DataLeads_Status['Причина отказа'] != "Пропала потребность/неактуально")) &
                                    (DataLeads_Status['На кого обращение'] != "Агапов Д.Г.") & ((DataLeads_Status['Должность'] == "Оператор колл-центра") | (DataLeads_Status['Должность'] == "Администратор колл-центра")) & (DataLeads_Status['Статус'] == 'Перешёл в сделку'))].copy()
    DataLeads_Status_err_1['Описание ошибки'] = "У сконвертированной сделки пристуствует причина отказа"
    DataLeads_Status_err_2 = DataLeads_Status[((DataLeads_Status['Причина отказа'].isnull()) & ((DataLeads_Status['Должность'] == "Оператор колл-центра") | (DataLeads_Status['Должность'] == "Администратор колл-центра")) & (DataLeads_Status['Статус'] == 'Забракован'))].copy()
    DataLeads_Status_err_2['Описание ошибки'] = "У забракованной сделки отсутствует причина отказа"
    DataLeads_Status_err_3 = DataLeads_Status[(((DataLeads_Status['Источник рекламы'].isnull()) | (DataLeads_Status['Источник лида'].isnull())) & (DataLeads_Status['Статус'] != 'Забракован') &         # У лида не заполнен источник или источник рекламы
                                    (DataLeads_Status['На кого обращение'] != "Агапов Д.Г.") & ((DataLeads_Status['Должность'] == "Оператор колл-центра") | (DataLeads_Status['Должность'] == "Администратор колл-центра")))].copy()
    DataLeads_Status_err_3['Описание ошибки'] = "У лида не заполнен источник или источник рекламы"
    DataLeads_Status_err_4 = DataLeads_Status[(DataLeads_Status['Причина отказа'].notnull()) & (DataLeads_Status['Статус'] != 'Забракован') & (DataLeads_Status['Статус'] != 'Перешёл в сделку') &       # У лида в процессе установлена причина отказа
                                    (DataLeads_Status['На кого обращение'] != "Агапов Д.Г.") & ((DataLeads_Status['Должность'] == "Оператор колл-центра") | (DataLeads_Status['Должность'] == "Администратор колл-центра"))].copy()
    DataLeads_Status_err_4['Описание ошибки'] = "У лида в процессе работы уже установлена причина отказа - необходимо установить статус 'Забракован'"
    DataLeads_Status_err_5 = DataLeads_Status[(DataLeads_Status['Отделение'].isnull()) & (DataLeads_Status['Статус'] != 'Забракован') &                                                                  # Не указано отделение в работе
                                    (DataLeads_Status['На кого обращение'] != "Агапов Д.Г.") & ((DataLeads_Status['Должность'] == "Оператор колл-центра") | (DataLeads_Status['Должность'] == "Администратор колл-центра"))].copy()
    DataLeads_Status_err_5['Описание ошибки'] = "Не указано отделение в работе"
    RESULT_err = pd.concat([RESULT_err, DataLeads_Status_err_1], axis=0)
    RESULT_err = pd.concat([RESULT_err, DataLeads_Status_err_2], axis=0)
    RESULT_err = pd.concat([RESULT_err, DataLeads_Status_err_3], axis=0)
    RESULT_err = pd.concat([RESULT_err, DataLeads_Status_err_4], axis=0)
    RESULT_err = pd.concat([RESULT_err, DataLeads_Status_err_5], axis=0)
    # Работа с новыми лидами и лидами в процессе
    DataLeads_Status_NEW = DataLeads_Status[((DataLeads_Status['Статус'] != 'Забракован') & (DataLeads_Status['Статус'] != 'Перешёл в сделку') &                                                         # Лиды в процессе работы
                                             ((DataLeads_Status['На кого обращение'] != "Агапов Д.Г.") & (DataLeads_Status['На кого обращение'] != "Христенко А.А.")) & ((DataLeads_Status['Должность'] == "Оператор колл-центра") | (DataLeads_Status['Должность'] == "Администратор колл-центра")))]
    # Фильтрация данных (не целевые лиды)
    DataLeadsFilter_Dont = DataLeads_Status[(((DataLeads_Status['Причина отказа'].notnull()) & (DataLeads_Status['Причина отказа'] != "Выбрал другую клинику/доктора") &
                                  (DataLeads_Status['Причина отказа'] != "Дорого/ не устраивает уровень цен") & (DataLeads_Status['Причина отказа'] != "Не отвечает на прогревы") &
                                  (DataLeads_Status['Причина отказа'] != "Отказался от дальнейшей связи") & (DataLeads_Status['Причина отказа'] != "Пропала потребность/неактуально")) &
                                  ((DataLeads_Status['Должность'] == "Оператор колл-центра") | (DataLeads_Status['Должность'] == "Администратор колл-центра") | (DataLeads_Status['Источник лида'] == "Instagram Христенко А.А.")))]
    # Фильтрация данных (целевые лиды)
    DataLeadsFilter = DataLeads_Status[((DataLeads_Status['Причина отказа'].isnull() | (DataLeads_Status['Причина отказа'] == "Выбрал другую клинику/доктора") |
                                  (DataLeads_Status['Причина отказа'] == "Дорого/ не устраивает уровень цен") | (DataLeads_Status['Причина отказа'] == "Не отвечает на прогревы") |
                                  (DataLeads_Status['Причина отказа'] == "Отказался от дальнейшей связи")  | (DataLeads_Status['Причина отказа'] == "Пропала потребность/неактуально")) &
                                  ((DataLeads_Status['Источник лида'] != "Instagram Христенко А.А.") & ((DataLeads_Status['Должность'] == "Оператор колл-центра") | (DataLeads_Status['Должность'] == "Администратор колл-центра"))))]

    DataLeadsFilter_call = DataLeadsFilter[(DataLeadsFilter['Источник лида'] == 'Звонок') | (DataLeadsFilter['Источник лида'] == 'Веб-сайт DEGA (запись с сайта)')]
    DataLeadsFilter_message = DataLeadsFilter[(DataLeadsFilter['Источник лида'] != 'Звонок') & (DataLeadsFilter['Источник лида'] != 'Веб-сайт DEGA (запись с сайта)')]
    # Нэйминг и исправление порядка столбцов
    RESULT_all = pd.concat([RESULT_all, DataLeads], axis=0)                                                                                                                                              # Собирем все периоды в общий фрейм
    RESULT_filter = pd.concat([RESULT_filter, DataLeadsFilter], axis=0)                                                                                                                                  # Собирем все периоды в общий фрейм
    RESULT_filter_call = pd.concat([RESULT_filter_call, DataLeadsFilter_call], axis=0)                                                                                                                   # Собирем все периоды в общий фрейм
    RESULT_filter_message = pd.concat([RESULT_filter_message, DataLeadsFilter_message], axis=0)                                                                                                          # Собирем все периоды в общий фрейм

Input_lag()
DataListFirst = pd.date_range(min(start_date, end_date), max(start_date, end_date)).strftime('%Y-%m-%dT%H:%M:%S').tolist()                                                                               # Создание списка дат
DataListLast = pd.date_range(min((start_date + timedelta(days=1)), (end_date + timedelta(days=1))), max((start_date), (end_date))).strftime('%Y-%m-%dT%H:%M:%S').tolist()                                # Создание списка дат
#Запуск цикла парсинга bitrix24
for datefirst, datelast in zip(DataListFirst, DataListLast):
    print("Выбранный диапазон дат: " + str(datefirst) + ", " + str(datelast))
    BitrixFrame(datefirst, datelast)
try:
    os.mkdir("Конверсия КЦ " + datename); os.mkdir("Конверсия КЦ " + datename + "/Инфографика")                                                                                                          # Создаём директории для хранения файлов
    Conv_all = Convertions(RESULT_filter)                                                                                                                                                                # Фрейм для конверсии КЦ
    Conv_call = Convertions(RESULT_filter_call)                                                                                                                                                          # Фрейм для конверсии звонков
    Conv_message = Convertions(RESULT_filter_message)                                                                                                                                                    # Фрейм для конверсии сообщений
    #Conv_source = Convertions_source(RESULT_filter)
    Make_plot()
    #SendTelegram("try", "")
except requests.exceptions.RequestException as err:
    print(err)
    #SendTelegram("except1", err)

# Запись в файл и редактор excel
with pd.ExcelWriter("Конверсия КЦ " + datename + '/Расчёты конверсии ' + str(datename) + '.xlsx', engine ="xlsxwriter") as writer:
    Conv_all.to_excel(writer, sheet_name='Конверсия по лидам', index=False, freeze_panes=(1,0))
    Conv_call.to_excel(writer, sheet_name='Конверсия по звонкам', index=False, freeze_panes=(1,0))
    Conv_message.to_excel(writer, sheet_name='Конверсия по сообщениям', index=False, freeze_panes=(1,0))
    #.to_excel(writer, sheet_name='Источники', index=False, freeze_panes=(1,0))
    #.to_excel(writer, sheet_name='Источники рекламы', index=False, freeze_panes=(1,0))
    DataLeads_Status_NEW.to_excel(writer, sheet_name='Лиды в работе', index=False, freeze_panes=(1,0))
    RESULT_err.to_excel(writer, sheet_name='Ошибки', index=False, freeze_panes=(1,0))
    RESULT_filter_call.to_excel(writer, sheet_name='Сконвертированные звонки', index=False, freeze_panes=(1,0))
    RESULT_filter_message.to_excel(writer, sheet_name='Сконвертированные переписки', index=False, freeze_panes=(1,0))
    RESULT_filter.to_excel(writer, sheet_name='Целевые для KPI', index=False, freeze_panes=(1,0))
    RESULT_all.to_excel(writer, sheet_name='Все лиды за период', index=False, freeze_panes=(1,0))

    workbook = writer.book                                                                                                                                                                               # Доступ к объектам (форматам, диаграммам xlsxwriter)
    # Лист Конверсия по всем лидам
    percent = workbook.add_format({"num_format" : "#,#0.0%"})                                                                                                                                            # Настройка формата отображения %
    header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top'})                                                                                                              # Формат для заголовков
    sheet = writer.sheets["Конверсия по лидам"]                                                                                                                                                          # Выбор активного листа
    sheet.autofilter('A1:E' + str(Conv_all.shape[0]))                                                                                                                                                    # Установка автофильтра на лист
    sheet.set_column(0, 0, 32); sheet.set_column(3, 3, 10); sheet.set_column(4, 4, 23)                                                                                                                   # Настройка ширины столбцов
    sheet.set_row(0, 45)                                                                                                                                                                                 # Настройка высоты строки
    sheet.set_column("D:D", 15, percent)                                                                                                                                                                 # Присвоение формата колонке
    sheet.conditional_format('D2:D100', {'type': '3_color_scale', 'mid_value': 70, 'min_value': 30})                                                                                                     # Настройка разброса градиента условного форматирования колонки
    for col_num, value in enumerate(Conv_all.columns.values):                                                                                                                                            # Присвоение форматирования заголовку
        sheet.write(0, col_num - 1 + 1, value, header_format)
    sheet.insert_image('G2', "Конверсия КЦ " + datename + "/Инфографика/Конверсия лидов по отделам.png")                                                                                                 # Вставка изображения
    sheet.insert_image('G33', "Конверсия КЦ " + datename + "/Инфографика/Конверсия лидов по операторам.png")                                                                                             # Вставка изображения

    # Лист Конверсия по звонкам
    percent = workbook.add_format({"num_format" : "#,#0.0%"})                                                                                                                                            # Настройка формата отображения %
    header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top'})                                                                                                              # Формат для заголовков
    sheet = writer.sheets["Конверсия по звонкам"]                                                                                                                                                        # Выбор активного листа
    sheet.autofilter('A1:E' + str(Conv_call.shape[0]))                                                                                                                                                   # Установка автофильтра на лист
    sheet.set_column(0, 0, 32); sheet.set_column(3, 3, 10); sheet.set_column(4, 4, 23)                                                                                                                   # Настройка ширины столбцов
    sheet.set_row(0, 45)                                                                                                                                                                                 # Настройка высоты строки
    sheet.set_column("D:D", 15, percent)                                                                                                                                                                 # Присвоение формата колонке
    sheet.conditional_format('D2:D100', {'type': '3_color_scale', 'mid_value': 70, 'min_value': 30})                                                                                                     # Настройка разброса градиента условного форматирования колонки
    for col_num, value in enumerate(Conv_call.columns.values):                                                                                                                                           # Присвоение форматирования заголовку
        sheet.write(0, col_num - 1 + 1, value, header_format)
    sheet.insert_image('G2', "Конверсия КЦ " + datename + "/Инфографика/Конверсия звонков по отделам.png")                                                                                               # Вставка изображения
    sheet.insert_image('G33', "Конверсия КЦ " + datename + "/Инфографика/Конверсия звонков по операторам.png")                                                                                           # Вставка изображения

    # Лист Конверсия по сообщениям
    percent = workbook.add_format({"num_format" : "#,#0.0%"})                                                                                                                                            # Настройка формата отображения %
    header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top'})                                                                                                              # Формат для заголовков
    sheet = writer.sheets["Конверсия по сообщениям"]                                                                                                                                                     # Выбор активного листа
    sheet.autofilter('A1:E' + str(Conv_message.shape[0]))                                                                                                                                                # Установка автофильтра на лист
    sheet.set_column(0, 0, 32); sheet.set_column(3, 3, 10); sheet.set_column(4, 4, 23)                                                                                                                   # Настройка ширины столбцов
    sheet.set_row(0, 45)                                                                                                                                                                                 # Настройка высоты строки
    sheet.set_column("D:D", 15, percent)                                                                                                                                                                 # Присвоение формата колонке
    sheet.conditional_format('D2:D100', {'type': '3_color_scale', 'mid_value': 70, 'min_value': 30})                                                                                                     # Настройка разброса градиента условного форматирования колонки
    for col_num, value in enumerate(Conv_message.columns.values):                                                                                                                                        # Присвоение форматирования заголовку
        sheet.write(0, col_num - 1 + 1, value, header_format)
    sheet.insert_image('G2', "Конверсия КЦ " + datename + "/Инфографика/Конверсия сообщений по отделам.png")                                                                                             # Вставка изображения
    sheet.insert_image('G33', "Конверсия КЦ " + datename + "/Инфографика/Конверсия сообщений по операторам.png")                                                                                         # Вставка изображения

    # Лист Лиды в работе
    sheet = writer.sheets['Лиды в работе']
    sheet.autofilter('A1:M' + str(DataLeads_Status_NEW.shape[0]))                                                                                                                                        # Установка автофильтра на лист

    # Лист Ошибки
    sheet = writer.sheets['Ошибки']
    sheet.autofilter('A1:N' + str(RESULT_err.shape[0]))                                                                                                                                                  # Установка автофильтра на лист
    sheet.insert_image('T2', "Конверсия КЦ " + datename + "/Инфографика/Распределение логических ошибок по причинам.png")                                                                                # Вставка изображения
    sheet.insert_image('T33', "Конверсия КЦ " + datename + "/Инфографика/Распределение логических ошибок по операторам.png")                                                                             # Вставка изображения

    # Лист Сконвертированные звонки
    sheet = writer.sheets['Сконвертированные звонки']
    sheet.autofilter('A1:M' + str(RESULT_filter_call.shape[0]))                                                                                                                                          # Установка автофильтра на лист

    # Лист Сконвертированные переписки
    sheet = writer.sheets['Сконвертированные переписки']
    sheet.autofilter('A1:M' + str(RESULT_filter_message.shape[0]))                                                                                                                                       # Установка автофильтра на лист

    # Лист Целевые для KPI
    sheet = writer.sheets['Целевые для KPI']
    sheet.autofilter('A1:M' + str(RESULT_filter.shape[0]))                                                                                                                                               # Установка автофильтра на лист

    # Лист Все лиды за период
    sheet = writer.sheets['Все лиды за период']
    sheet.autofilter('A1:L' + str(RESULT_all.shape[0]))                                                                                                                                                  # Установка автофильтра на лист


















































'''
# Выгрузка дополнительных полей
#deals = pd.DataFrame(bitra.get_all('crm.deal.list', params={'select': ['*', 'UF_*'], 'filter': {'>DATE_CREATE': str(datefirst), '<DATE_CREATE': str(datelast)}})) # Список лидов
#leads = pd.DataFrame(bitra.get_all('crm.lead.list', params={'select': ['*', 'UF_*'], 'filter': {'>DATE_CREATE': str(datefirst), '<DATE_CREATE': str(datelast)}})) # Список сделок
#users = pd.DataFrame(bitra.get_all('user.get')) # Получить список пользователей
status = pd.DataFrame(bitra.get_all('status.get')) # Получить список статусов crm.status.list()
#tittles = pd.DataFrame(tittles) # Получить описание полей
with pd.ExcelWriter('Описание полей ' + str(datetime.now().strftime('%H-%M-%S')) + '.xlsx') as writer:
    status.to_excel(writer, sheet_name='Мдааа')
'''
'''
contacts = pd.DataFrame(bitra.get_all('crm.contact.fields'))                                                                                                                                             # Получение описания полей контактов
def ContactID(ID):                                                                                                                                                                                       # Получение расширенного списка полей для лида
    contact = bitra.get_by_ID('crm.contact.get', [ID])['PHONE']
    print(contact)
    return contact
ContactID(129092)
'''
'''
Conv_call = pd.pivot_table(RESULT_filter_call,
               index=["Оператор", "Отделение"],
               columns=["Статус"],
               values=["Все обращения"],
               aggfunc=[len, lambda x: len(x[(x == 'Перешёл в сделку')])],
               fill_value=0, margins=True)
#Conv_call = pd.pivot_table(RESULT_filter_call, index=['Отделение', 'Оператор'], columns=['Статус'], values=['Статус'], aggfunc={"Quantity": lambda x: len(x[(x == 'Перешёл в сделку')]), "Price": 'count'}, fill_value=0, margins=True)
print(Conv_call)
#Base_operator.loc[Base_operator['% конверсии'].notna(), '% конверсии'] = pd.Series(["{0:.2f}%".format(val * 100) for val in Base_operator['% конверсии']],index = Base_operator.index)
#["{0:.2f}%".format(val * 100) for val in df["Diff%"]], index = df.index
#Base_operator.loc[Base_operator['% конверсии'].notna(), '% конверсии'] = Base_operator['% конверсии'].map(lambda n: '{:.1%}'.format(n))                                                             # Изменение формата отображения % конверсий
'''