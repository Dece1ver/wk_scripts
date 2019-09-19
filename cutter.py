# -*- coding: utf-8 -*-

import os
import sys
import tkinter as tk
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
start_table = PrettyTable()
lines = []


max_weigth = 480

ver = 0.0919

os.system('cls')
print('Если параметр RETRACT_FEED не равен 0, то разделение может быть некорректным!')
print('# # # # # # # # # # # # # # # # # # # # # # # #\nВыбрать управляющую программу, чтобы разделить.')
try:
    file_path = askopenfilename()
    win32gui.SetForegroundWindow(window)
    os.system('cls')
    print(f'Ver.{ver}')
    with open(file_path, 'r') as file:
        state = 1
        try:
            if file.read(2) not in ('%\n', 'O0'):
                state = 0
        except UnicodeDecodeError:
            state = 0
        if not state:
            print('Файл не подходит.')
        else:
            #
            # старт
            #
            file.seek(0)
            print(f'Файл {file_path} открыт, чтение...')
            weigth = round(os.path.getsize(file_path) / 1024)
            if weigth <= max_weigth:
                print(f'\nОбъем файла: {weigth:,} / {max_weigth} Кб.'.replace(',', ' '))
                print('Файл норм, ничего не надо резать.')
            else:
                parts = weigth // max_weigth + 1
                for line in file:
                    lines.append(line.strip())
                    print(f'Строк: {len(lines)}\r', end='')

                for line in lines:
                    if line.find('G01') != -1:
                        if line.find('F') != -1:
                            feed = line[line.find('F'):].strip()
                            break

                if not feed:
                    feed = int(input(('Подача не обнаружена, ввести подачу без точки: ')))
                    feed = f'F{feed}.'

                start_table.field_names = ['File', 'Info']
                start_table.add_row(['Объем файла', f'{weigth:,} Кб.'.replace(',', ' ')])
                start_table.add_row(['Выбран максимальный объем', f'{max_weigth} Кб'])
                start_table.add_row(['Количество частей', f'{parts}'])
                start_table.add_row(['Количество строк', len(lines)])
                abt_strings = str(len(lines) // parts)
                start_table.add_row(['Строк в каждой программе', f'≈ {abt_strings[:-3]}k'])
                start_table.add_row(['Стартовая подача', feed])
                print(start_table)
                #
                # шапка
                #
                head_line = 0
                while lines[head_line].find('G01') == -1:
                    head_line += 1

                head_table = PrettyTable()
                head_table.field_names = ['№', 'Строка']
                head_table.align['№'] = 'l'
                head_table.align['Строка'] = 'l'
                head = lines[:head_line]
                head_title = '\n# # # # # Авто-шапка # # # # #'
                print(head_title)
                for n, i in enumerate(head):
                    head_table.add_row([n + 1, i])
                print(head_table)
                head_override = input('\nЕсли шапка норм, то Enter, если нет, то ввести номер строки конца шапки: ')
                while head_override != '' and head_override.isdigit():
                    head = lines[:int(head_override)]
                    os.system('cls')
                    head_table.clear_rows()
                    print(f'Ver.{ver}')
                    print(f'Файл {file_path} открыт, чтение...')
                    print(f'Строк: {len(lines)}\r')
                    print(start_table)
                    head_title = '\n# # # # # # Шапка # # # # # # - ИЗМЕНЕНА'
                    print(head_title)
                    for n, i in enumerate(head):
                        head_table.add_row([n + 1, i])
                    print(head_table)
                    head_override = input('\nЕсли шапка норм, то Enter, если нет, то ввести номер строки конца шапки: ')
                #
                # концовка
                #
                end_line = -1
                while lines[end_line].find('G00') == -1:
                    end_line -= 1

                end_table = PrettyTable()
                end_table.field_names = ['№', 'Строка']
                end_table.align['№'] = 'l'
                end_table.align['Строка'] = 'l'
                end = lines[end_line:]
                print('\n# # # # # Авто-концовка # # # # #')
                for n, i in enumerate(end):
                    end_table.add_row([(len(lines) + end_line + n - 1), i])
                print(end_table)
                end_line_override = input('\nЕсли концовка норм, то Enter, если нет, то ввести номер строки с которой начать: ')
                while end_line_override != '' and end_line_override.strip('-').isdigit():
                    end_line_override = int(end_line_override) + 1
                    end_line = end_line_override
                    end = lines[end_line:]
                    os.system('cls')
                    end_table.clear_rows()
                    print(f'Ver.{ver}')
                    print(f'Файл {file_path} открыт, чтение...')
                    print(f'Строк: {len(lines)}\r')
                    print(start_table)
                    print(head_title)
                    print(head_table)
                    print('\n# # # # # Концовка # # # # # - ИЗМЕНЕНА')
                    for n, i in enumerate(end):
                        end_table.add_row([(end_line + n - 1), i])
                    print(end_table)
                    end_line_override = input('\nЕсли концовка норм, то Enter, если нет, то ввести номер строки с которой начать: ')
                #
                # разделение
                #
                lines = lines[head_line:end_line]
                cut_lines = round(len(lines) / parts)
                start_line = 0
                last_line = start_line + cut_lines
                print(f'Общее количество строк без шапки и конца: {len(lines)}')
                print(f'Среднее количество строк в каждой программе: {cut_lines}')
                for part in range(parts):
                    print(f'\n# # # # # # # # # # #\nСоставление части: {part+1}')
                    if part > 0:
                        print(f'Координаты XY в шапке изменены на: {last_xy}')
                    print(f'Начальная строка: ({start_line + len(head) + 1}) {lines[start_line]}')
                    if part + 1 == parts:
                        print('Последний кусок.')
                        with open(f'{file_path}-{part + 1}', 'w') as part_file:
                            name_line = 0
                            for i, line in enumerate(head):
                                if line.find(')') != -1:
                                    name_line = i
                                    break
                            if name_line != 0:
                                print(f'Обнаружено имя: ({name_line + 1}) {head[name_line]}', end='')
                                old_name = head[name_line]
                                head[name_line] = head[name_line].replace(r')', f'-{part + 1})')
                                print(f' -> {head[name_line]}')
                            part_file.write('\n'.join(head[:-1]) + '\n')

                            feed_line = 0
                            for i, line in enumerate(lines[start_line:last_line]):
                                if line.find('G01') != -1:
                                    feed_line = start_line + i
                                    break
                            print(f'Найдена строка для вставки подачи: ({feed_line+len(head)}) {lines[feed_line]}', end='')
                            lines[feed_line] = lines[feed_line] + feed
                            print(f' -> {lines[feed_line]}')
                            part_file.write('\n'.join(lines[start_line:last_line]) + '\n')
                            part_file.write('\n'.join(end))
                            print(f'Файл {file_path}-{part + 1} записан.')
                            print('Конец.')
                            break
                    print(f'Поиск со строки: ({last_line + len(head) + 1}) {lines[last_line]}')
                    for n, line in enumerate(lines[last_line:]):
                        if line.find('G00') == -1:
                            if line.find('X') != -1 and line.find('Y') != -1:
                                last_xy_index = n + last_line
                            continue
                        else:
                            last_line = n + last_line
                            print(f'Обнаружено место для разделения: ({last_line + len(head) + 1}) {line}')
                            with open(f'{file_path}-{part + 1}', 'w') as part_file:
                                name_line = 0
                                for i, line in enumerate(head):
                                    if line.find(')') != -1:
                                        name_line = i
                                        break
                                if name_line != 0:
                                    print(f'Обнаружено имя: ({name_line + 1}) {head[name_line]}', end='')
                                    old_name = head[name_line]
                                    head[name_line] = head[name_line].replace(r')', f'-{part + 1})')
                                    print(f' -> {head[name_line]}')
                                part_file.write('\n'.join(head if part == 0 else head[:-1]) + '\n')

                                for i, line in enumerate(head):
                                    if line.find('X') != -1 and line.find('Y') != -1:
                                        last_xy = lines[last_xy_index].replace('G01', 'G00').replace('G02', 'G00').replace('G03', 'G00')
                                        if last_xy.find('G00') == -1:
                                            ind = last_xy.find('X')
                                            last_xy = f'G00{last_xy[ind:]}'
                                        head[i] = last_xy
                                if part > 0:
                                    feed_line = 0
                                    for i, line in enumerate(lines[start_line: last_line]):
                                        if line.find('G01') != -1:
                                            feed_line = start_line + i
                                            break

                                    print(f'Найдена строка для вставки подачи: ({feed_line+len(head)}) {lines[feed_line]}', end='')
                                    lines[feed_line] = lines[feed_line] + feed
                                    print(f' -> {lines[feed_line]}')
                                    part_file.write('\n'.join(lines[start_line: last_line]) + '\n')
                                    part_file.write('\n'.join(end))
                                else:
                                    part_file.write('\n'.join(lines[start_line: last_line]) + '\n')
                                    part_file.write('\n'.join(end))
                                if name_line != 0:
                                    head[name_line] = old_name
                                print(f'Файл {file_path}-{part+1} записан.')
                            start_line = last_line
                            last_line = start_line + cut_lines
                            break
except FileNotFoundError:
    print('Отмена выбора файла')
input('Enter для выхода.')
