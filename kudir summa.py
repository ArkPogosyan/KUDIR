import pandas as pd
import os


def read_and_process_file(file_path):
    """
    Читает Excel-файл и возвращает DataFrame с данными.
    """
    try:
        # Чтение файла Excel
        df = pd.read_excel(file_path, engine='openpyxl')

        # Приводим названия столбцов к стандартному виду
        df.columns = [
            col.strip().lower()
            .replace(' ', '_')
            .replace('-', '_')
            .replace('"', '')
            .replace("'", '')
            for col in df.columns
        ]

        # Выводим названия столбцов для диагностики
        print(f"Названия столбцов в файле '{file_path}': {df.columns.tolist()}")

        # Определяем обязательные столбцы
        required_columns = ['контрагент', 'всего_расходов', 'для_налоговой_базы']

        # Создаем словарь соответствия реальных названий столбцов и ожидаемых
        column_mapping = {
            'контрагент': None,
            'всего_расходов': None,
            'для_налоговой_базы': None
        }

        # Находим ближайшие совпадения для каждого обязательного столбца
        for col in df.columns:
            if 'контрагент' in col.lower():
                column_mapping['контрагент'] = col
            elif 'всего' in col.lower() and 'расход' in col.lower():
                column_mapping['всего_расходов'] = col
            elif 'налог' in col.lower() and 'баз' in col.lower():
                column_mapping['для_налоговой_базы'] = col

        # Переименовываем столбцы в DataFrame
        df.rename(columns=column_mapping, inplace=True)

        # Проверяем наличие всех обязательных столбцов
        missing_cols = [col for col, mapped_col in column_mapping.items() if mapped_col is None]
        if missing_cols:
            print(f"Ошибка: В файле '{file_path}' отсутствуют обязательные столбцы: {', '.join(missing_cols)}")
            return None

        return df
    except Exception as e:
        print(f"Ошибка чтения файла '{file_path}': {str(e)}")
        return None


def merge_data(files):
    """
    Объединяет данные из всех файлов в один DataFrame.
    """
    all_data = []

    for file in files:
        print(f"Обработка файла: {file}")
        df = read_and_process_file(file)
        if df is not None:
            all_data.append(df)

    if not all_data:
        print("Нет данных для обработки.")
        return None

    # Объединяем все DataFrame в один
    merged_df = pd.concat(all_data, ignore_index=True)

    # Группируем данные по контрагентам и суммируем значения
    grouped_data = merged_df.groupby('контрагент', as_index=False).agg({
        'всего_расходов': 'sum',
        'для_налоговой_базы': 'sum'
    })

    # Вычисляем разницу между "Всего расходов" и "Для налоговой базы"
    grouped_data['разница'] = grouped_data['всего_расходов'] - grouped_data['для_налоговой_базы']

    return grouped_data


def save_results_to_excel(data, output_file):
    """
    Сохраняет результаты в Excel-файл.
    """
    try:
        data.to_excel(output_file, index=False, engine='openpyxl')
        print(f"Итоговые результаты сохранены в файл: {output_file}")
    except Exception as e:
        print(f"Ошибка сохранения файла: {str(e)}")


# Основная программа
if __name__ == "__main__":
    # Запрашиваем пути к входным файлам
    input_files = []
    for i in range(1, 5):  # Предполагается 4 файла
        file_path = input(f"Введите путь к файлу {i}: ").strip()
        if not os.path.isfile(file_path):
            print(f"Ошибка: Файл '{file_path}' не существует.")
            exit()
        input_files.append(file_path)

    # Запрашиваем путь для выходного файла
    output_file = input(
        "\nВведите имя файла для сохранения итоговых результатов (например, total_results.xlsx): ").strip()
    if not output_file.endswith('.xlsx'):
        output_file += '.xlsx'

    # Обработка данных
    merged_data = merge_data(input_files)
    if merged_data is not None:
        save_results_to_excel(merged_data, output_file)