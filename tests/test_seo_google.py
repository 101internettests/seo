from pages.data import query_new, query_new_spb, query_new_msk, query_new_ekb, query_new_ufa, query_new_krasnodar, query_new_novosibirsk
import time
from selenium.webdriver.common.by import By
import pandas as pd


def test_find_competitors_and_website_rank_spb(driver):
    websites_to_check = ["piter-online"]
    results = []

    for website_to_check in websites_to_check:
        for query in query_new_spb:
            driver.get(f"https://www.google.com/search?q={query}")
            time.sleep(2)  # Небольшая пауза для загрузки страницы
            search_results = driver.find_elements(By.CSS_SELECTOR, "div.g")

            website_rank = None
            link = None
            top_three_links = []  # Список, чтобы хранить топ-3 ссылки

            for index, result in enumerate(search_results, start=1):
                if len(top_three_links) >= 3:  # Если уже найдены топ-3 ссылки
                    break
                link_elements = result.find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    url = link_elements[0].get_attribute("href")  # Получаем URL
                    if url not in top_three_links:  # Проверяем на уникальность
                        top_three_links.append(url)  # Добавляем ссылку в топ-3

            # Заполняем недостающие ссылки "нет ссылки"
            while len(top_three_links) < 3:
                top_three_links.append('нет ссылки')

            # Поиск ранга для нужного сайта
            for index, result in enumerate(search_results, start=1):
                link_elements = result.find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    url = link_elements[0].get_attribute("href")
                    if url and any(website in url for website in websites_to_check):
                        website_rank = index
                        link = url
                        break

            # Формируем строку результата
            if website_rank is not None and link is not None:
                result_row = [website_to_check, query, website_rank, link] + top_three_links
            else:
                result_row = [website_to_check, query, 'не найдено', 'нет ссылки'] + top_three_links

            results.append(result_row)  # Добавляем строку результата
            time.sleep(2)  # Задержка между запросами

    # Сохраняем результаты в файл Excel
    df = pd.DataFrame(results, columns=['Сайт', 'Запрос', 'Место в поиске', 'Ссылка', 'Ссылка1', 'Ссылка2', 'Ссылка3'])

    try:
        df.to_excel('google_spb.xlsx', index=False)  # Пытаемся сохранить файл
        print("Результаты успешно сохранены в google_spb.xlsx.")
    except Exception as e:
        print(f"Не удалось сохранить результаты: {e}")


def test_find_competitors_and_website_rank_msk(driver):
    websites_to_check = ["moskvaonline"]
    results = []

    for website_to_check in websites_to_check:
        for query in query_new_msk:
            driver.get(f"https://www.google.com/search?q={query}")
            time.sleep(2)  # Небольшая пауза для загрузки страницы
            search_results = driver.find_elements(By.CSS_SELECTOR, "div.g")

            website_rank = None
            link = None
            top_three_links = []  # Список, чтобы хранить топ-3 ссылки

            for index, result in enumerate(search_results, start=1):
                if len(top_three_links) >= 3:  # Если уже найдены топ-3 ссылки
                    break
                link_elements = result.find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    url = link_elements[0].get_attribute("href")  # Получаем URL
                    if url not in top_three_links:  # Проверяем на уникальность
                        top_three_links.append(url)  # Добавляем ссылку в топ-3

            # Заполняем недостающие ссылки "нет ссылки"
            while len(top_three_links) < 3:
                top_three_links.append('нет ссылки')

            # Поиск ранга для нужного сайта
            for index, result in enumerate(search_results, start=1):
                link_elements = result.find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    url = link_elements[0].get_attribute("href")
                    if url and any(website in url for website in websites_to_check):
                        website_rank = index
                        link = url
                        break

            # Формируем строку результата
            if website_rank is not None and link is not None:
                result_row = [website_to_check, query, website_rank, link] + top_three_links
            else:
                result_row = [website_to_check, query, 'не найдено', 'нет ссылки'] + top_three_links

            results.append(result_row)  # Добавляем строку результата
            time.sleep(2)  # Задержка между запросами

    # Сохраняем результаты в файл Excel
    df = pd.DataFrame(results, columns=['Сайт', 'Запрос', 'Место в поиске', 'Ссылка', 'Ссылка1', 'Ссылка2', 'Ссылка3'])

    try:
        df.to_excel('google_msk.xlsx', index=False)  # Пытаемся сохранить файл
        print("Результаты успешно сохранены в google_msk.xlsx.")
    except Exception as e:
        print(f"Не удалось сохранить результаты: {e}")


def test_find_competitors_and_website_rank_ufa(driver):
    websites_to_check = ["101internet"]
    results = []

    for website_to_check in websites_to_check:
        for query in query_new_ufa:
            driver.get(f"https://www.google.com/search?q={query}")
            time.sleep(2)  # Небольшая пауза для загрузки страницы
            search_results = driver.find_elements(By.CSS_SELECTOR, "div.g")

            website_rank = None
            link = None
            top_three_links = []  # Список, чтобы хранить топ-3 ссылки

            for index, result in enumerate(search_results, start=1):
                if len(top_three_links) >= 3:  # Если уже найдены топ-3 ссылки
                    break
                link_elements = result.find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    url = link_elements[0].get_attribute("href")  # Получаем URL
                    if url not in top_three_links:  # Проверяем на уникальность
                        top_three_links.append(url)  # Добавляем ссылку в топ-3

            # Заполняем недостающие ссылки "нет ссылки"
            while len(top_three_links) < 3:
                top_three_links.append('нет ссылки')

            # Поиск ранга для нужного сайта
            for index, result in enumerate(search_results, start=1):
                link_elements = result.find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    url = link_elements[0].get_attribute("href")
                    if url and any(website in url for website in websites_to_check):
                        website_rank = index
                        link = url
                        break

            # Формируем строку результата
            if website_rank is not None and link is not None:
                result_row = [website_to_check, query, website_rank, link] + top_three_links
            else:
                result_row = [website_to_check, query, 'не найдено', 'нет ссылки'] + top_three_links

            results.append(result_row)  # Добавляем строку результата
            time.sleep(2)  # Задержка между запросами

    # Сохраняем результаты в файл Excel
    df = pd.DataFrame(results, columns=['Сайт', 'Запрос', 'Место в поиске', 'Ссылка', 'Ссылка1', 'Ссылка2', 'Ссылка3'])

    try:
        df.to_excel('google_ufa.xlsx', index=False)  # Пытаемся сохранить файл
        print("Результаты успешно сохранены в google_ufa.xlsx.")
    except Exception as e:
        print(f"Не удалось сохранить результаты: {e}")


def test_find_competitors_and_website_rank_kransnodar(driver):
    websites_to_check = ["101internet"]
    results = []

    for website_to_check in websites_to_check:
        for query in query_new_krasnodar:
            driver.get(f"https://www.google.com/search?q={query}")
            time.sleep(2)  # Небольшая пауза для загрузки страницы
            search_results = driver.find_elements(By.CSS_SELECTOR, "div.g")

            website_rank = None
            link = None
            top_three_links = []  # Список, чтобы хранить топ-3 ссылки

            for index, result in enumerate(search_results, start=1):
                if len(top_three_links) >= 3:  # Если уже найдены топ-3 ссылки
                    break
                link_elements = result.find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    url = link_elements[0].get_attribute("href")  # Получаем URL
                    if url not in top_three_links:  # Проверяем на уникальность
                        top_three_links.append(url)  # Добавляем ссылку в топ-3

            # Заполняем недостающие ссылки "нет ссылки"
            while len(top_three_links) < 3:
                top_three_links.append('нет ссылки')

            # Поиск ранга для нужного сайта
            for index, result in enumerate(search_results, start=1):
                link_elements = result.find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    url = link_elements[0].get_attribute("href")
                    if url and any(website in url for website in websites_to_check):
                        website_rank = index
                        link = url
                        break

            # Формируем строку результата
            if website_rank is not None and link is not None:
                result_row = [website_to_check, query, website_rank, link] + top_three_links
            else:
                result_row = [website_to_check, query, 'не найдено', 'нет ссылки'] + top_three_links

            results.append(result_row)  # Добавляем строку результата
            time.sleep(2)  # Задержка между запросами

    # Сохраняем результаты в файл Excel
    df = pd.DataFrame(results, columns=['Сайт', 'Запрос', 'Место в поиске', 'Ссылка', 'Ссылка1', 'Ссылка2', 'Ссылка3'])

    try:
        df.to_excel('google_kransnodar.xlsx', index=False)  # Пытаемся сохранить файл
        print("Результаты успешно сохранены в google_kransnodar.xlsx.")
    except Exception as e:
        print(f"Не удалось сохранить результаты: {e}")


def test_find_competitors_and_website_rank_novosibirsk(driver):
    websites_to_check = ["101internet"]
    results = []

    for website_to_check in websites_to_check:
        for query in query_new_novosibirsk:
            driver.get(f"https://www.google.com/search?q={query}")
            time.sleep(2)  # Небольшая пауза для загрузки страницы
            search_results = driver.find_elements(By.CSS_SELECTOR, "div.g")

            website_rank = None
            link = None
            top_three_links = []  # Список, чтобы хранить топ-3 ссылки

            for index, result in enumerate(search_results, start=1):
                if len(top_three_links) >= 3:  # Если уже найдены топ-3 ссылки
                    break
                link_elements = result.find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    url = link_elements[0].get_attribute("href")  # Получаем URL
                    if url not in top_three_links:  # Проверяем на уникальность
                        top_three_links.append(url)  # Добавляем ссылку в топ-3

            # Заполняем недостающие ссылки "нет ссылки"
            while len(top_three_links) < 3:
                top_three_links.append('нет ссылки')

            # Поиск ранга для нужного сайта
            for index, result in enumerate(search_results, start=1):
                link_elements = result.find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    url = link_elements[0].get_attribute("href")
                    if url and any(website in url for website in websites_to_check):
                        website_rank = index
                        link = url
                        break

            # Формируем строку результата
            if website_rank is not None and link is not None:
                result_row = [website_to_check, query, website_rank, link] + top_three_links
            else:
                result_row = [website_to_check, query, 'не найдено', 'нет ссылки'] + top_three_links

            results.append(result_row)  # Добавляем строку результата
            time.sleep(2)  # Задержка между запросами

    # Сохраняем результаты в файл Excel
    df = pd.DataFrame(results, columns=['Сайт', 'Запрос', 'Место в поиске', 'Ссылка', 'Ссылка1', 'Ссылка2', 'Ссылка3'])

    try:
        df.to_excel('google_novosibirsk.xlsx', index=False)  # Пытаемся сохранить файл
        print("Результаты успешно сохранены в google_novosibirsk.xlsx.")
    except Exception as e:
        print(f"Не удалось сохранить результаты: {e}")