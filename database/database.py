"""Модуль отвечает за работу с базой данных"""

import sqlite3
import os


def create_database():
    """Функция создает базу данных"""

    conn = sqlite3.connect('example1.db')
    c = conn.cursor()

    c.execute('''CREATE TABLE IF NOT EXISTS my_table
                 (col1 TEXT, col2 TEXT, col3 TEXT, col4 BOOLEAN)''')

    conn.commit()

    conn.close()


def fill_data(list1, list2, string):
    """Функция заполняет строку в базе данных."""

    conn = sqlite3.connect('example1.db')
    c = conn.cursor()

    for item1, item2 in zip(list1, list2):
        if item1 is None:
            col4_value = "False"
        else:
            col4_value = "True"
        if item1 is not None:
            item1 = os.path.splitext(os.path.basename(item1))[0]
        if item2 is not None:
            item2 = os.path.splitext(os.path.basename(item2))[0]
        c.execute("INSERT INTO my_table VALUES (?, ?, ?, ?)", (item1, item2, string, col4_value))

    conn.commit()

    conn.close()


def view_data():
    """Функция выводит в консоль содержимое базы данных."""

    conn = sqlite3.connect('example.db')
    c = conn.cursor()

    c.execute("SELECT * FROM my_table")

    rows = c.fetchall()

    for row in rows:
        print(row)

    conn.close()
