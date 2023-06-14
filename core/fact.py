"""Модуль отвечает за внесение факта в project."""

import datetime
import logging
import os
import time

import pandas as pd
import pythoncom
import win32com.client as win32

import settings.readOF as config


def get_excel_pd(path):
    """Создает DataFrame из обменной формы"""

    if not os.path.isabs(path):
        logging.warning('%s: Путь до ОФ не абсолютный', get_excel_pd.__name__)
    logging.info('%s: Пытаемся записать ОФ в DataFrame', get_excel_pd.__name__)
    try:
        data = pd.read_excel(path)
    except Exception:
        logging.error('%s: Не получилось записать ОФ в DataFrame', get_project.__name__)
        raise Exception('Не получилось записать ОФ в DataFrame')
    logging.info('%s: ОФ успешно записалась в DataFrame', get_excel_pd.__name__)
    return data

def get_project(path):
    """Открывает файл проекта и возвращает объект проекта.

    На вход поступает путь до файла проекта.
    На выход объект project и MS Project.
    """
    if not os.path.isabs(path):
        logging.warning('%s: Путь до файла проекта не абсолютный', get_project.__name__)
    logging.info('%s: Пытаемся открыть файл проекта', get_project.__name__)
    try:
        msp = win32.Dispatch("MSProject.Application", pythoncom.CoInitialize())
        _abs_path = os.path.abspath(path)
        print(_abs_path)
        msp.FileOpen(_abs_path)
        project = msp.ActiveProject
    except Exception:
        logging.error('%s: Файл проекта не смог открыться', get_project.__name__)
        raise Exception('Не получилось открыть файл проекта')
    logging.info('%s: Файл проекта успешно открылся', get_project.__name__)
    return project, msp


def _get_data_task(t):
    """"Получает значения task из нужных столбцов

    На вход поступает объект Task из project. Из него достаются
    требуемые данные, ими заполняется список, который возвращается
    для дальнейшего использования.
    """
    arr = []
    try:
        for i in config.ID_COLUMN.keys():
            data = getattr(t,i)
            if isinstance(data, datetime.datetime):
                data = datetime.datetime.date(data)
            arr.append(data)
    except Exception as e:
        print(e)
        logging.error('%s: Неверный идентификатор столбца project', _get_data_task.__name__)
        raise Exception('Неверный идентификатор столбца project')
    return arr


def fill_dataframe(project):
    """Заполняет DataFrame значениями из project

    На вход поступают объект project. Формируется dataframe
    с данными из требуемых столбцов и возвращается
    для дальнейшего использования.
    """
    logging.info('%s: Создаем DataFrame из столбцов объекта проекта', fill_dataframe.__name__)
    if not project:
        logging.error('%s: Не удалось получить объект проекта', fill_dataframe.__name__)
        raise Exception("Объект проекта пустой")
    if not config.ID_COLUMN:
        logging.error('%s: Ключевые столбцы не заданы', fill_dataframe.__name__)
        raise Exception("Ключевые столбцы не заданы")
    task_collection = project.Tasks
    data = pd.DataFrame(columns=config.ID_COLUMN.values())
    start = time.time()
    try:
        for t in task_collection:
            data.loc[len(data.index)] = _get_data_task(t)
    except Exception:
        logging.error('%s: Неверно заполнен словарь столбцов и их идентификаторов', fill_dataframe.__name__)
        raise Exception("Ошибка в словаре слобцов и их идентификаторов")
    logging.info('%s: DataFrame из столбцов объекта проекта успешно создан', fill_dataframe.__name__)

    return data

def _check_str(excel_str, project_str, columns):
    """Сравнивает значения task ОФ и проекта

    На вход поступает строка из Dataframe, полученного из excel и строка
    из Dataframe, полученного из project, а также колонки для сравнения.
    Функция сравнивает ячейки в переданных столбцах по очереди и возвращает
    False и название столбца, в котором найдено несоответствие или
    True и None, если несоответствий нет
    """
    for col in columns:
        if not pd.isnull(excel_str[col]) and excel_str[col] != 'НД':
            excel_str[col] = datetime.datetime.date(excel_str[col])
        if excel_str[col] != project_str[col]:
            return False, col
    return True, None


def check_form(data_project, data_excel, columns):
    """Находит несоответствия между обменной формой и проектом и сохраняет их в словарь

    На вход поступает 2 объекта Dataframe и список столбцов,
    по которым будем происходить сравнение. Функция находит несоответствия
    и заполняет словарь в котором хранится имя столбца, где найдено
    несоответствие, и значение из обменной формы. В качестве ключа индекс
    строки, в которой найдено несоответствие.
    """
    logging.info('%s: Ищем несоответсвия между ОФ и проектом', check_form.__name__)
    if data_project.empty:
        logging.error('%s: DataFrame проекта пустой', check_form.__name__)
        raise Exception('DataFrame проекта пустой')
    if data_excel.empty:
        logging.error('%s: DataFrame ОФ пустой', check_form.__name__)
        raise Exception('DataFrame ОФ пустой')
    changes = {}
    for i, excel_str in data_excel.iterrows():
        for j, project_str in data_project.iterrows():
            if excel_str[columns[0]] == project_str[columns[0]]:
                status, column = _check_str(excel_str, project_str, columns[1:])
                if not status:
                    changes[j] = [column, excel_str[column]]
                break
    logging.info('%s: Поиск несоответсвия между ОФ и проектом окончен', check_form.__name__)
    return changes

def change_project(project, msp, changes):
    """Вносит изменения в проект.

    На вход поступает объект project и объект приложения MS Project,
    а также словарь изменений. Функция проходит по project построчно и
    ищет индексы, которые совпадают со значениями ключей. Затем
    вносит изменения в соответствующий столбец соответствующей строки
    project.
    """
    if not project:
        logging.info('%s: Объект проекта пустой', change_project.__name__)
        raise Exception('Объект проекта пустой')
    task_collection = project.Tasks
    if not changes:
        logging.info('%s: Изменений в проекте нет', change_project.__name__)
    else:
        logging.info('%s: Применяем изменения', change_project.__name__)
        try:
            for i, t in enumerate(task_collection):
                if i in changes.keys():
                    #Не получается присвоить значение в t.ActualStart.
                    #t.ActualStart = t.ActualStart.replace(day=changes[i][1])
                    #pywintypes.datetime и datetime.date
                    new_date = datetime.datetime(changes[i][1].year, changes[i][1].month, changes[i][1].day)
                    t.ActualStart = new_date
                    print(t.Active)
                    t.Name = t.Name.replace('п', '1')
            msp.FileSave()
        except Exception as e:
            logging.error('%s: Не получилось применить изменения', change_project.__name__)
            raise Exception('Не получилось применить изменения')
        logging.info('%s: Изменения успешно применены', change_project.__name__)


def main(path_to_project, path_to_excel):
    """Управляющая функция"""

    try:
        excel_df = get_excel_pd(path_to_excel)
        project, msp = get_project(path_to_project)
        project_df = fill_dataframe(project)
        results = check_form(project_df, excel_df, "УИД_(П)")
        change_project(project, msp, results)
    except Exception as e:
        print("smth bad")


