"""Модуль отвечает за проверки доступности папок и копирование данных."""

import os
import shutil


def check_folder_writable(folder_path):
    """Функция проверяет папку на доступность для записи в нее."""

    return os.access(folder_path, os.W_OK)


def check_folder_readable(file_paths):
    """Функция проверяет папку на доступность для чтения."""

    for file_path in file_paths:
        if not os.access(file_path, os.R_OK):
            return False
    return True


def transfer_files(file_paths, destination_folder):
    """Копирует файлы из одной папки в другую.

    На вход поступают абсолютные пути до файлов, которые нужно скопировать
    и путь до папки, в которую нужно скопировать.
    """
    for file_path in file_paths:
        if file_path is not None:
            try:
                file_name = os.path.basename(file_path)
                destination_file = os.path.join(destination_folder, file_name)
                shutil.copyfile(file_path, destination_file)
                print(f"Файл {file_name} успешно загружен в папку {destination_folder}.")
            except Exception as e:
                return e
    return True

