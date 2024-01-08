from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl

# Считываем артикулы из файла wb.txt
with open("wb.txt", "r") as file:
    articles = file.read().splitlines()

# Создать объект Options для Chrome
options = Options()
options.add_argument("--headless")  # Для запуска без отображения браузера
options.add_argument("--log-level=3")

# Инициализация драйвера с использованием Options
with webdriver.Chrome(options=options) as driver:
    # Создать файл Excel и рабочий лист
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # Уменьшить время ожидания до 3 секунд
    wait = WebDriverWait(driver, 3)

    try:
        for article in articles:
            # Формируем URL с использованием артикула
            url = f"https://www.wildberries.ru/catalog/{article}/detail.aspx"

            # Перейти на страницу
            driver.get(url)

            # Ждем загрузки элемента
            try:
                # Проверяем наличие элемента с ценой
                price_element = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".price-block__final-price")))
            except:
                # Обработка случая, когда цена не найдена (например, "Нет в наличии")
                sheet.append([f"{article}", "Нет в наличии"])
                print(f"Товар с артикулом {article} отсутствует в наличии.")
                continue  # пропустить текущую итерацию, так как товар отсутствует

            # Получить текст из элемента
            price_text = price_element.text.strip()

            # Оставить только цифры
            price_digits = ''.join(filter(str.isdigit, price_text))

            # Получить название товара
            name_element = driver.find_element(By.CLASS_NAME, "product-page__title")
            name = name_element.text.strip()

            # Добавить значения в строку таблицы
            row_data = [article, price_digits, name]
            sheet.append(row_data)

            # Вывести цену в консоль
            print(f"Цена успешно записана в файл Excel для артикула {article}: {price_digits} руб - {name}")

    finally:
        # Сохранить Excel файл
        workbook.save("output.xlsx")

        print(f"Работа парсера завершена!")