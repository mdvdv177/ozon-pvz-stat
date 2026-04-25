import csv
import sys
import os
import tempfile
import pandas as pd
from collections import Counter, defaultdict

def get_main_cell(cells):
    """
    Возвращает номер ячейки, в которой больше всего позиций.
    При равенстве — с меньшим номером.
    """
    counts = Counter(cells)
    # Находим максимальное количество
    max_count = max(counts.values())
    # Собираем ячейки с максимальным количеством
    candidates = [cell for cell, cnt in counts.items() if cnt == max_count]
    # Сортируем как числа (если возможно) или как строки
    try:
        candidates.sort(key=int)
    except ValueError:
        candidates.sort()
    return candidates[0]

def process_inventory_file(file_path):
    """
    Обрабатывает xlsx-файл с остатками на складе и выводит заказы с несколькими позициями.
    """
    orders = defaultdict(list)  # заказ -> список ячеек
    
    df = pd.read_excel(file_path, skiprows=5, usecols='B:H')
    df.to_csv(f'{os.path.dirname(os.path.abspath(__file__))}/result.csv', index=False, encoding='utf-8')
    
    with open(f'{os.path.dirname(os.path.abspath(__file__))}/result.csv', 'r', encoding='utf-8-sig') as f:
        # Пропускаем начальные строки до заголовка
        header_found = False
        header_row = None
        for line in f:
            if 'Номер отправления' in line and 'Ячейка' in line:
                header_found = True
                header_row = line
                break
        if not header_found:
            print("Не найден заголовок с колонками 'Номер отправления' и 'Ячейка'")
            return

        # Определяем индексы нужных столбцов
        reader = csv.reader([header_row])
        headers = next(reader)
        try:
            track_num_idx = headers.index('Номер отправления')
            cell_idx = headers.index('Ячейка')
        except ValueError as e:
            print(f"Ошибка: не найдена колонка {e}")
            return

        # Читаем остальные строки данных
        csv_reader = csv.reader(f)
        for row in csv_reader:
            if len(row) <= max(track_num_idx, cell_idx):
                continue  # пропускаем пустые строки
            track_num = row[track_num_idx].strip()
            cell_raw = row[cell_idx].strip()
            if not track_num or not cell_raw:
                continue

            # Извлекаем идентификатор заказа (часть до первого '-') 
            order_id = track_num.split('-')[0]

            # Извлекаем номер стеллажа (часть до первого '-')
            cell = cell_raw.split('-')[0]
            orders[order_id].append(cell)

    # Фильтруем заказы с количеством позиций > 1
    multi_item_orders = {oid: cells for oid, cells in orders.items() if len(cells) > 1}
    if not multi_item_orders:
        print("Нет заказов с несколькими позициями.")
        return

    # Сортируем заказы по основной ячейке (где больше всего позиций)
    def sort_key(item):
        order_id, cells = item
        main_cell = get_main_cell(cells)
        # Преобразуем в число, если возможно, для корректной сортировки
        try:
            return int(main_cell)
        except ValueError:
            return main_cell

    sorted_orders = sorted(multi_item_orders.items(), key=sort_key)

    # Выводим результат
    out = """
        <h1 style='font-size: 48px;'>Отчет об обработке</h1>
        <hr>
        <p style='font-size: 30px; white-space: pre-wrap;'>"""
    for i, (order_id, cells) in enumerate(sorted_orders, 1):
        total = len(cells)
        cell_counts = Counter(cells)
        cell_parts = []
        for cell, count in cell_counts.items():
            cell_parts.append(f"яч <i>{cell}</i> - {count}")
        cell_str = ", ".join(cell_parts)
        plus = 0
        if i>=100:
            plus+=2
        elif i>=10:
            plus+=1
        if total>=10:
            plus+=1
        output = f"{i} ({total}){' '*(3-plus)}: {cell_str}"
        if total>3:
            output = f"<b>{output}</b>"
        out += output + '<br>'
        
    # Создаем временный файл, который сам удалится позже
    with tempfile.NamedTemporaryFile('w', suffix='.html', delete=False) as f:
        f.write(out+'</p>')
        path = f.name

# Открываем в браузере или приложении по умолчанию
    os.system(f"open {path}")
    
if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Использование: python(или python3) script.py <путь к xlsx файлу>")
        sys.exit(1)
    process_inventory_file(sys.argv[1])
