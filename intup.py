# -*- coding: utf-8 -*-

import os
import sys
import tkinter as tk
import re
import time
from tkinter.filedialog import askopenfilename
try:
    import win32gui
except ImportError:
    ask_for_install_ptable = input("Не установлена библиотека PyWin32, попробовать установить? (+ или Y - согласие, пропуск - отказ) \n")
    if ask_for_install_ptable in ("+", "y", "Y"):
        try:
            install_cmd = "pip install pywin32 pywin32-ctypes"
            print(install_cmd)
            os.system(install_cmd)
            os.startfile(__file__)
            sys.exit(0)
        except Exception as e:
            print(e)
    else:
        input("Отмена. Необходимо установить вручную (pip install pywin32)")
        sys.exit(0)

try:
    from prettytable import PrettyTable
except ImportError:
    ask_for_install_ptable = input("Не установлена библиотека PrettyTable, попробовать установить? (+ или Y - согласие, пропуск - отказ)\n")
    if ask_for_install_ptable in ("+", "y", "Y"):
        try:
            install_cmd = "pip install PTable"
            print(install_cmd)
            os.system(install_cmd)
            from prettytable import PrettyTable
        except Exception as e:
            print(e)
    else:
        input("Отмена. Необходимо установить вручную (pip install PTable)")
        sys.exit(0)

window = win32gui.GetForegroundWindow()
root = tk.Tk()
root.withdraw()
new_lines = []
state = 0
max_old = 0
max_new = 0


table = PrettyTable()
table.field_names = ["До", "После", "Подробно"]
table.align['До'] = 'l'
table.align['После'] = 'l'


def windowEnumerationHandler(hwnd, top_windows):
    top_windows.append((hwnd, win32gui.GetWindowText(hwnd)))


top_windows = []
win32gui.EnumWindows(windowEnumerationHandler, top_windows)


def set_max_len(line, max_len):
    if len(line) > max_len:
        max_len = len(line)
    return max_len


def printer(string, timing=0):
    for i in string:
        print(i, end="", flush=True)
        time.sleep(timing)
    print()


os.system('cls')
print(r"""
     ______          __
    /\__  _\        /\ \__
    \/_/\ \/     ___\ \ ,_\    __     __   _ __    __   __  _
       \ \ \   /' _ `\ \ \/  /'__`\ /'_ `\/\`'__\/'__`\/\ \/'\
        \_\ \__/\ \/\ \ \ \_/\  __//\ \L\ \ \ \//\  __/\/>  </
        /\_____\ \_\ \_\ \__\ \____\ \____ \ \_\\ \____\/\_/\_\
        \/_____/\/_/\/_/\/__/\/____/\/___L\ \/_/ \/____/\//\/_/
                       ____           /\____/             __
                      /\  _`\         \_/__/   __        /\ \__
                      \ \,\L\_\    ___   _ __ /\_\  _____\ \ ,_\
                       \/_\__ \   /'___\/\`'__\/\ \/\ '__`\ \ \/
                         /\ \L\ \/\ \__/\ \ \/ \ \ \ \ \L\ \ \ \_
                         \ `\____\ \____\\ \_\  \ \_\ \ ,__/\ \__\
                          \/_____/\/____/ \/_/   \/_/\ \ \/  \/__/
                                                      \ \_\
                                                       \/_/ @dece1ver
""")

try:
    file_path = askopenfilename()
    win32gui.SetForegroundWindow(window)
    os.system("cls")
    if file_path == "":
        raise Exception('Выбор файла отменен.')
    elif not file_path.endswith(".eia") and not file_path.endswith(".EIA"):
        raise Exception("Выбрана не УП Mazatrol")
    with open(file_path, "r") as file:
        printer(f"\n| Открыт файл {file_path}\n| Добавить отвод по Z перед поворотом по оси B на 90°? (+ или Y - согласие, пропуск - отказ)")
        add_z = input("| ")
        lines = file.readlines()
        for line in lines:
            if re.findall("T\d+", line) and not line.find("T#20") != -1:
                new_line = re.sub("T\d+", "T#20", line, 1)
                table.add_row([line.rstrip("\n"), new_line.rstrip("\n"), "Номер инструмента изменен на макропеременную"])
                max_old = set_max_len(line, max_old)
                max_new = set_max_len(new_line, max_new)
                line = new_line
                state += 1
            if add_z in ("+", "y", "Y") and line.find("B90.") != -1:
                new_line = f"G0 Z#26\n{line}"
                table.add_row(["\n" + line.rstrip("\n"), new_line.rstrip("\n"), "ВНИМАНИЕ!\nДобавлен отвод по Z с установкой макропеременной"])
                max_old = set_max_len(line, max_old)
                max_new = set_max_len(new_line, max_new)
                add_z = 0
                line = new_line
                state += 1
            if line.find("G54") != -1 and not line.find("(G54)") != -1:
                new_line = line.replace("G54", "(G54)")
                table.add_row([line.rstrip("\n"), new_line.rstrip("\n"), "G54 закомментирован"])
                max_old = set_max_len(line, max_old)
                max_new = set_max_len(new_line, max_new)
                line = new_line
                state += 1
            if line.find("M30") != -1:
                new_line = line.replace("M30", "M99")
                table.add_row([line.rstrip("\n"), new_line.rstrip("\n"), "Конец подпрограммы установлен"])
                max_old = set_max_len(line, max_old)
                max_new = set_max_len(new_line, max_new)
                line = new_line
                state += 1
            new_lines.append(line)
    new_file_name = file_path.replace(".eia", "_r.eia")
    if state > 0:
        os.system(f"mode con cols={max_old + max_new + 53} lines=30")
        add_z_res = "будет" if add_z in ("+", "y", "Y", 0) else "не будет"
        print(f"\n| Открыт файл {file_path}\n| Отвод {add_z_res} добавлен")
        print(table.get_string(title="Информация"))
        with open(new_file_name, "w") as new_file:
            new_file.write("\n".rstrip("\n").join(new_lines))
            printer(f"\n| Файл записан в {new_file_name}\n| Enter для открытия программы...")
        start_cmd = input("| ")
        if start_cmd.lower() in ("-n", "--notepad", "--блокнот"):
            printer("| Запуск блокнотом")
            os.system(f"notepad.exe {new_file_name}")
        elif start_cmd.lower() in ("-t", "--type"):
            new_file_name = new_file_name.replace("/", "\\")
            printer('| Вывод УП в консоль:\n')
            os.system(f'type "{new_file_name}"')
            input()
        elif start_cmd.lower() in ("-ts", "--typeandstart"):
            new_file_name = new_file_name.replace("/", "\\")
            printer('| Вывод УП в консоль с последующим запуском:\n')
            os.system(f'type "{new_file_name}"')
            input()
            printer("| Запуск программой по умолчанию")
            os.system(new_file_name)
        else:
            printer("| Запуск программой по умолчанию")
            os.system(new_file_name)
    else:
        input("| Файл не был изменен, т.к. не было найдено заменяемых позиций.")
except Exception as e:
    input(f"Исключение: {e} \nТип: {type(e)}")
