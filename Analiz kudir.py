import pandas as pd
import re
import os


def read_excel_file(file_path):
    """
    Читает Excel-файл и возвращает DataFrame.
    """
    try:
        # Чтение файла Excel, пропуск второй строки с номерами столбцов
        df = pd.read_excel(file_path, skiprows=[1], engine='openpyxl')
        return df
    except Exception as e:
        print(f"Ошибка чтения файла: {str(e)}")
        return None


def extract_contragent(text):
    """
    Извлекает название контрагента из текста.
    Признаки контрагента: ООО, АО, ПАО, ИП перед названием в кавычках.
    """
    pattern = r'(ООО|АО|ПАО|ИП)\s+"([^"]+)"'
    match = re.search(pattern, text, re.IGNORECASE)
    if match:
        return f"{match.group(1).upper()} \"{match.group(2)}\""
    return None


def process_data(df):
    """
    Обрабатывает данные, находит всех контрагентов и вычисляет их параметры.
    """
    contragents_data = {}

    for _, row in df.iterrows():
        operation = str(row['содержание_операции'])
        contragent = extract_contragent(operation)

        if contragent:
            total_expenses = pd.to_numeric(
                str(row['расходы_-_всего']).replace(',', '.').replace(' ', ''),
                errors='coerce'
            )
            tax_base_expenses = pd.to_numeric(
                str(row['в_т.ч._расходы,_учитываемые_при_исчислении_налоговой_базы']).replace(',', '.').replace(' ',
                                                                                                                ''),
                errors='coerce'
            )

            if contragent not in contragents_data:
                contragents_data[contragent] = {
                    'total_expenses': 0,
                    'tax_base_expenses': 0
                }

            contragents_data[contragent]['total_expenses'] += total_expenses if pd.notnull(total_expenses) else 0
            contragents_data[contragent]['tax_base_expenses'] += tax_base_expenses if pd.notnull(
                tax_base_expenses) else 0

    return contragents_data


def save_results_to_excel(contragents_data, output_file):
    """
    Сохраняет результаты в Excel-файл.
    """
    results = []
    for contragent, data in contragents_data.items():
        results.append({
            'Контрагент': contragent,
            'Всего расходов': data['total_expenses'],
            'Для налоговой базы': data['tax_base_expenses'],
            'Разница': data['total_expenses'] - data['tax_base_expenses']
        })

    result_df = pd.DataFrame(results)
    result_df.to_excel(output_file, index=False, engine='openpyxl')
    print(f"Результаты сохранены в файл: {output_file}")


# Основная программа
if __name__ == "__main__":
    # Предлагаем путь по умолчанию
    default_path = r"C:\Users\Ark\Downloads\data.xlsx"
    input_file = input(f"Введите путь к входному Excel-файлу (по умолчанию: {default_path}): ").strip()

    # Если пользователь не ввел путь, используем путь по умолчанию
    if not input_file:
        input_file = default_path

    # Проверяем существование файла
    if not os.path.isfile(input_file):
        print("Ошибка: Указанный файл не существует.")
        exit()

    # Чтение файла
    df = read_excel_file(input_file)
    if df is None:
        exit()

    # Приводим названия столбцов к стандартному виду
    df.columns = [col.strip().lower().replace(' ', '_') for col in df.columns]

    # Проверяем наличие нужных столбцов
    required_columns = [
        'содержание_операции',
        'расходы_-_всего',
        'в_т.ч._расходы,_учитываемые_при_исчислении_налоговой_базы'
    ]

    missing_cols = [col for col in required_columns if col not in df.columns]
    if missing_cols:
        print(f"Ошибка: Отсутствуют обязательные столбцы: {', '.join(missing_cols)}")
        exit()

    # Обработка данных
    contragents_data = process_data(df)

    # Запрашиваем имя выходного файла
    output_file = input("\nВведите имя файла для сохранения результатов (например, results.xlsx): ").strip()
    if not output_file.endswith('.xlsx'):
        output_file += '.xlsx'

    # Сохраняем результаты
    save_results_to_excel(contragents_data, output_file)