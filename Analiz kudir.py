import pandas as pd
import re
from pathlib import Path
import tkinter as tk
from tkinter import filedialog


def extract_contractor(description):
    """
    Функция для извлечения названия контрагента из описания.
    """
    if not isinstance(description, str):
        return None

    # Регулярное выражение для поиска контрагентов
    patterns = [
        r'"(?:ООО|АО|ИП|ПАО)\s+([^"]+)"',  # Например: "ООО АМД КОМПАНИ"
        r'(?:ООО|АО|ИП|ПАО)\s+"?([^"]+)"?',  # Например: ООО "АТЛАНТИДА" или ООО АТЛАНТИДА
        r'"ТВ\s+"[^"]+"\s+И\s+КО"',  # Например: "ТВ "АД РУСС" И КО"
    ]

    for pattern in patterns:
        match = re.search(pattern, description)
        if match:
            contractor = match.group(1).strip()
            return contractor

    return None


def process_excel_file(file_path):
    """
    Основная функция для обработки файла Excel.
    """
    try:
        # Чтение файла, пропуск первых 6 строк
        df = pd.read_excel(file_path, header=None, skiprows=6)

        # Переименование столбцов для удобства работы
        df.columns = ['ID', 'Date_and_Doc', 'Description', 'Contractor_Info', 'Income', 'Expenses']

        # Удаление строк, где столбец Expenses пустой или NaN
        df = df.dropna(subset=['Expenses'])

        # Очистка столбца Expenses от запятых и преобразование в числа
        df['Expenses'] = df['Expenses'].apply(
            lambda x: float(str(x).replace(',', '.')) if isinstance(x, str) else x
        )

        # Извлечение названий контрагентов из столбца Contractor_Info
        df['Contractor'] = df['Contractor_Info'].apply(extract_contractor)

        # Удаление строк без найденных контрагентов
        df = df.dropna(subset=['Contractor'])

        # Группировка данных по контрагентам и суммирование расходов
        result = df.groupby('Contractor')['Expenses'].sum().reset_index()

        # Округление сумм до двух знаков после запятой
        result['Expenses'] = result['Expenses'].round(2)

        # Путь для сохранения результата
        output_path = file_path.parent / f"result_{file_path.stem}.xlsx"

        # Сохранение результата в новый файл
        result.to_excel(output_path, index=False, engine='openpyxl')
        print(f"Результат сохранен: {output_path}")

    except Exception as e:
        print(f"Ошибка при обработке файла: {e}")


def main():
    """
    Главная функция программы.
    """
    # Создаем скрытое окно Tkinter
    root = tk.Tk()
    root.withdraw()

    # Открываем диалог выбора файла
    file_path = filedialog.askopenfilename(
        title="Выберите файл Excel",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )

    if not file_path:
        print("Файл не выбран. Программа завершена.")
        return

    # Обрабатываем выбранный файл
    process_excel_file(Path(file_path))


if __name__ == "__main__":
    main()