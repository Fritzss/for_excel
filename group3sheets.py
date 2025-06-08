from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from os import getcwd as pwd, remove as rm
from functools import wraps
import time
import configparser
import chardet
from copy import copy


print('начинаю, читаю конфиг')
config = configparser.ConfigParser()


def detect_encoding(fp):
    with open(fp, 'rb') as f:
        det = chardet.universaldetector.UniversalDetector()
        for l in f:
            det.feed(l)
            if det.done:
                break
        det.close()
        return det.result['encoding']


config.read('.\\config.ini', encoding=detect_encoding('.\\config.ini'))

src_file = config.get('Settings', 'src_file')
sheet = config.get('Settings', 'sheet')
tag = config.get('Settings', 'tag')

print(f'читаю файл {src_file}, исходный лист {sheet}')
workdir = pwd()
wb = load_workbook(f'{workdir}\\{src_file}', read_only=False, keep_vba=False)
ws = wb[sheet]
target_path = f'{workdir}\\groups2sheets-{src_file}'


def timeit(func):
    @wraps(func)
    def timeit_wrapper(*args, **kwargs):
        start_time = time.perf_counter()
        result = func(*args, **kwargs)
        end_time = time.perf_counter()
        total_time = end_time - start_time
        print(f'Function {func.__name__} time {total_time:.4f} seconds')
        return result

    return timeit_wrapper


@timeit
def first_row(ws):
    f_row = 0
    for row in ws.rows:
        if ws.row_dimensions[row[0].row].outlineLevel == 0 and ws.row_dimensions[row[0].row + 1].outlineLevel == 1:
            f_row = row[0].row
            break
    return f_row


def last_row(ws, tag):
    l_row = 0
    for row in ws.rows:
        if tag in str(row[0].value):
            l_row = row[0].row
            break
        if l_row == 0:
            l_row = ws.max_row
    return l_row


def print_phrase(n: int, word: str) -> None:
    # Определяем род слова (упрощенно)
    last_char = word[-1].lower()
    is_feminine = last_char in ('а')

    # Выбираем форму глагола
    if n == 1:
        verb = "создана" if is_feminine else "создан"
    else:
        verb = "создано"

    # Определяем форму слова
    n_abs = abs(n)
    if 11 <= (n_abs % 100) <= 14:
        form = 2
    else:
        remainder = n_abs % 10
        if remainder == 1:
            form = 0
        elif 2 <= remainder <= 4:
            form = 1
        else:
            form = 2

    # Склоняем слово
    endings = {
        'а': ['а', 'ы', ''],
        'я': ['я', 'и', 'й'],
        'м': ['', 'а', 'ов']  # для мужского рода
    }
    if is_feminine:
        key = 'а' if last_char == 'а' else 'м'
        base = word[:-1]
        word_form = base + endings[key][form]
    else:
        base = word
        if form == 0:
            word_form = base
        elif form == 1:
            word_form = base + 'а'
        else:
            word_form = base + 'ов'

    print(f"{verb} {n} {word_form}")


@timeit
def cp_header(wb, filials, target_file, last_header_row):
    # Исходный лист
    source_sheet = wb[sheet]

    # Собираем объединенные ячейки в строках 1-5
    merged_ranges = []
    for merged_range in source_sheet.merged_cells.ranges:
        if merged_range.min_row <= last_header_row and merged_range.max_row >= 1:
            merged_ranges.append(merged_range)

    # Список новых листов
    new_sheet_names = filials[1]

    for name in new_sheet_names:
        new_sheet = wb.create_sheet(title=name["filial"])

        # Копирование значений и стилей
        for row in range(1, (last_header_row)):
            for col in range(1, (last_header_row)):
                source_cell = source_sheet.cell(row=row, column=col)
                new_cell = new_sheet.cell(row=row, column=col)
                new_cell.value = source_cell.value
                # Исправленное копирование стилей
                if source_cell.has_style:
                    new_cell.font = copy(source_cell.font)
                    new_cell.border = copy(source_cell.border)
                    new_cell.fill = copy(source_cell.fill)
                    new_cell.number_format = source_cell.number_format
                    new_cell.alignment = copy(source_cell.alignment)
        # Восстановление объединенных ячеек
        for merged_range in merged_ranges:
            cell_range = f"{get_column_letter(merged_range.min_col)}{merged_range.min_row}:" \
                         f"{get_column_letter(merged_range.max_col)}{merged_range.max_row}"
            new_sheet.merge_cells(cell_range)
        # Копирование ширины столбцов
        for col in range(1, last_header_row):
            col_letter = get_column_letter(col)
            new_sheet.column_dimensions[col_letter].width = \
                source_sheet.column_dimensions[col_letter].width

        for row in range(1, (last_header_row)):
            new_sheet.row_dimensions[row].height = source_sheet.row_dimensions[row].height
    # Сохранение изменений
    wb.save(target_file)
    wb.close()


@timeit
def copy_rows_with_grouping(sheet, cities, target_file, last_header_row):
    """
    Копирует строки с группировками из исходного листа в новый лист
    """
    wb = load_workbook(target_file, read_only=False, keep_vba=False)
    src_ws = wb[sheet]
    filials = cities[1]
    all_sheets = len(filials)
    count = 0
    for target_sheet in filials:
        new_ws = wb[target_sheet['filial']]
        start_row = target_sheet['first_row']
        end_row = target_sheet['last_row']
        print(f'осталось {all_sheets - count}, копирую лист{target_sheet["filial"]}, строки c {start_row} по {end_row}')
        count = count + 1
#################################################
    # Копирование данных и стилей
        for row_idx in range(start_row, end_row + 1):
            for col_idx in range(1, src_ws.max_column + 1):
                src_cell = src_ws.cell(row=row_idx, column=col_idx)
              # print(f'row {row_idx} new {new_row_idx} start {start_row} new start {new_start_row}')
                new_cell = new_ws.cell(
                    row=row_idx - start_row + last_header_row,
                    column=col_idx,
                    value=src_cell.value
                )

                # Копирование стилей с использованием copy()
                if src_cell.has_style:
                    new_cell.font = copy(src_cell.font)
                    new_cell.border = copy(src_cell.border)
                    new_cell.fill = copy(src_cell.fill)
                    new_cell.number_format = src_cell.number_format
                    new_cell.alignment = copy(src_cell.alignment)

            # Копирование параметров строки
            new_row = row_idx - start_row + last_header_row
            new_ws.row_dimensions[new_row].outline_level = src_ws.row_dimensions[row_idx].outline_level
            new_ws.row_dimensions[new_row].height = src_ws.row_dimensions[row_idx].height


        # Копирование ширины столбцов
        for col_idx in range(1, src_ws.max_column + 1):
            col_letter = get_column_letter(col_idx)
            new_ws.column_dimensions[col_letter].width = \
                src_ws.column_dimensions[col_letter].width

    wb.save(target_file)
    wb.close()

@timeit
def get_cities(ws):
    l_row = last_row(ws, tag)
    filials = []
    for row in ws.rows:
        row = row[0].row
        if ws.row_dimensions[row].outlineLevel == 0 and ws.row_dimensions[row + 1].outlineLevel == 1:
            if len(filials) == 0:
                filials.append({"filial": f" {(list(ws.values)[row - 1])[0]}", "first_row": row})
            else:
                filials.append({"filial": f" {(list(ws.values)[row - 1])[0]}", "first_row": row})
                filials[len(filials) - 2].update({"last_row": row - 1})
        if row >= l_row - 1:
            filials[len(filials) - 1].update({"last_row": row - 1})
            break
    return l_row, filials


print(f'ищу группы на листе {sheet}')
cities = get_cities(ws)
print('создаю листы и копирую шапку')
last_header_row = first_row(ws)
cp_header(wb=wb,
          filials=cities,
          target_file=target_path,
          last_header_row=last_header_row)

copy_rows_with_grouping(sheet=sheet,
                        cities=cities,
                        target_file=target_path,
                        last_header_row=last_header_row)

print(f'новая книга сохранена в файл {target_path}')
input('press any key for close')
