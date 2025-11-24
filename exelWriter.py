import os
from openpyxl import Workbook # type: ignore

def main():
    # Создаём новый рабочий файл Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "User Data"

    # Заголовки
    ws.append(["Имя", "Возраст", "Email"])

    print("Введите данные для записи в Excel (оставьте имя пустым для завершения):")

    while True:
        name = input("Имя: ").strip()
        if not name:
            break
        try:
            age = int(input("Возраст: "))
        except ValueError:
            print("Некорректный возраст. Пропускаем запись.")
            continue
        email = input("Email: ").strip()

        # Добавляем строку в Excel
        ws.append([name, age, email])

    excel_filename = "user_data.xlsx"

    wb.save(excel_filename)
    print(f"\nДанные успешно сохранены в файл: {os.path.abspath(excel_filename)}")

if __name__ == "__main__":
    main()