import docx
import csv


def word_to_csv(word_file, csv_file):
    """
    Преобразует таблицу из файла Word в файл CSV.

    Аргументы:
        word_file (str): Путь к файлу Word.
        csv_file (str): Путь к файлу CSV.
    """

    doc = docx.Document(word_file)
    table = doc.tables[0]  # Предполагается, что таблица одна

    data = []
    for row in table.rows:
        row_data = []
        for cell in row.cells:
            row_data.append(cell.text)
        data.append(row_data)

    with open(csv_file, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerows(data)


# Пример использования
word_file = "input.docx"  # Замените на путь к вашему файлу Word
csv_file = "output.csv"  # Замените на путь к желаемому файлу CSV

word_to_csv(word_file, csv_file)
