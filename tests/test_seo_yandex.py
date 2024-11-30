from selenium.webdriver.common.by import By
import time
from pages.data_yandex import url_ms, url_spb, url_ufa
from pages.data_yandex_second import url_krasnodar, url_novosibirsk, url_kazan, url_samara
from pages.data_google_second import query_new_msk, query_new_spb, query_new_ufa, query_new_krasnodar
from pages.data_google import query_new_novosibirsk, query_new_kazan, query_new_samara
import pandas as pd


def test_find_competitors_and_website_rank_msk(driver):
    results = []
    websites_to_check = ["moskvaonline.ru"]

    for website_to_check in websites_to_check:
        for url, filename in zip(url_ms, query_new_msk):
            driver.get(url)
            search_results = driver.find_elements(By.CSS_SELECTOR, "li[data-cid]")
            top_three_links = []  # Store top three links
            website_rank = None
            link = None

            # Look for the website in the search results
            for result in search_results:
                link_elements = result.find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    url = link_elements[0].get_attribute("href")
                    if any(website in url for website in websites_to_check):
                        website_rank = search_results.index(result) + 1
                        link = url
                        break

            # Get the top three links
            for i in range(min(3, len(search_results))):
                link_elements = search_results[i].find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    top_three_links.append(link_elements[0].get_attribute("href"))

            # Ensure there are always 3 links
            top_three_links.extend([None] * (3 - len(top_three_links)))  # Pad with None if fewer than 3 links

            # If website was found, create a valid result row
            if website_rank is not None and link is not None:
                result_row = [website_to_check, filename, website_rank, link]
            else:
                result_row = [website_to_check, filename, 'не найдено', 'нет ссылки']

            result_row.extend(top_three_links)  # Add the top 3 links to the result
            results.append(result_row)

    # Convert to DataFrame and write to Excel
    df = pd.DataFrame(results,
                      columns=['Сайт', 'Запрос', 'Место в поиске', 'Ссылка', 'Ссылка1', 'Ссылка2', 'Ссылка3'])
    df.to_excel('yandex_msk.xlsx', index=False)


def test_find_competitors_and_website_rank_spb(driver):
    results = []
    websites_to_check = ["piter-online.net"]
    for website_to_check in websites_to_check:
        for url, filename in zip(url_spb, query_new_spb):
            driver.get(url)
            search_results = driver.find_elements(By.CSS_SELECTOR, "li[data-cid]")
            top_three_links = []  # Store top three links
            website_rank = None
            link = None

            # Look for the website in the search results
            for result in search_results:
                link_elements = result.find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    url = link_elements[0].get_attribute("href")
                    if any(website in url for website in websites_to_check):
                        website_rank = search_results.index(result) + 1
                        link = url
                        break

            # Get the top three links
            for i in range(min(3, len(search_results))):
                link_elements = search_results[i].find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    top_three_links.append(link_elements[0].get_attribute("href"))

            # Ensure there are always 3 links
            top_three_links.extend([None] * (3 - len(top_three_links)))  # Pad with None if fewer than 3 links

            # If website was found, create a valid result row
            if website_rank is not None and link is not None:
                result_row = [website_to_check, filename, website_rank, link]
            else:
                result_row = [website_to_check, filename, 'не найдено', 'нет ссылки']

            result_row.extend(top_three_links)  # Add the top 3 links to the result
            results.append(result_row)

    # Convert to DataFrame and write to Excel
    df = pd.DataFrame(results,
                      columns=['Сайт', 'Запрос', 'Место в поиске', 'Ссылка', 'Ссылка1', 'Ссылка2', 'Ссылка3'])
    df.to_excel('yandex_spb.xlsx', index=False)


def test_find_competitors_and_website_rank_ufa(driver):
    results = []
    websites_to_check = ["101internet.ru"]
    for website_to_check in websites_to_check:
        for url, filename in zip(url_ufa, query_new_ufa):
            driver.get(url)
            search_results = driver.find_elements(By.CSS_SELECTOR, "li[data-cid]")
            top_three_links = []  # Store top three links
            website_rank = None
            link = None

            # Look for the website in the search results
            for result in search_results:
                link_elements = result.find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    url = link_elements[0].get_attribute("href")
                    if any(website in url for website in websites_to_check):
                        website_rank = search_results.index(result) + 1
                        link = url
                        break

            # Get the top three links
            for i in range(min(3, len(search_results))):
                link_elements = search_results[i].find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    top_three_links.append(link_elements[0].get_attribute("href"))

            # Ensure there are always 3 links
            top_three_links.extend([None] * (3 - len(top_three_links)))  # Pad with None if fewer than 3 links

            # If website was found, create a valid result row
            if website_rank is not None and link is not None:
                result_row = [website_to_check, filename, website_rank, link]
            else:
                result_row = [website_to_check, filename, 'не найдено', 'нет ссылки']

            result_row.extend(top_three_links)  # Add the top 3 links to the result
            results.append(result_row)

    # Convert to DataFrame and write to Excel
    df = pd.DataFrame(results,
                      columns=['Сайт', 'Запрос', 'Место в поиске', 'Ссылка', 'Ссылка1', 'Ссылка2', 'Ссылка3'])
    df.to_excel('yandex_ufa.xlsx', index=False)


def test_find_competitors_and_website_rank_krasnodar(driver):
    results = []
    websites_to_check = ["101internet.ru"]
    for website_to_check in websites_to_check:
        for url, filename in zip(url_krasnodar, query_new_krasnodar):
            driver.get(url)
            search_results = driver.find_elements(By.CSS_SELECTOR, "li[data-cid]")
            top_three_links = []  # Store top three links
            website_rank = None
            link = None

            # Look for the website in the search results
            for result in search_results:
                link_elements = result.find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    url = link_elements[0].get_attribute("href")
                    if any(website in url for website in websites_to_check):
                        website_rank = search_results.index(result) + 1
                        link = url
                        break

            # Get the top three links
            for i in range(min(3, len(search_results))):
                link_elements = search_results[i].find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    top_three_links.append(link_elements[0].get_attribute("href"))

            # Ensure there are always 3 links
            top_three_links.extend([None] * (3 - len(top_three_links)))  # Pad with None if fewer than 3 links

            # If website was found, create a valid result row
            if website_rank is not None and link is not None:
                result_row = [website_to_check, filename, website_rank, link]
            else:
                result_row = [website_to_check, filename, 'не найдено', 'нет ссылки']

            result_row.extend(top_three_links)  # Add the top 3 links to the result
            results.append(result_row)

    # Convert to DataFrame and write to Excel
    df = pd.DataFrame(results,
                      columns=['Сайт', 'Запрос', 'Место в поиске', 'Ссылка', 'Ссылка1', 'Ссылка2', 'Ссылка3'])
    df.to_excel('yandex_krasnodar.xlsx', index=False)


def test_find_competitors_and_website_rank_novosibirsk(driver):
    results = []
    websites_to_check = ["101internet.ru"]
    for website_to_check in websites_to_check:
        for url, filename in zip(url_novosibirsk, query_new_novosibirsk):
            driver.get(url)
            search_results = driver.find_elements(By.CSS_SELECTOR, "li[data-cid]")
            top_three_links = []  # Store top three links
            website_rank = None
            link = None

            # Look for the website in the search results
            for result in search_results:
                link_elements = result.find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    url = link_elements[0].get_attribute("href")
                    if any(website in url for website in websites_to_check):
                        website_rank = search_results.index(result) + 1
                        link = url
                        break

            # Get the top three links
            for i in range(min(3, len(search_results))):
                link_elements = search_results[i].find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    top_three_links.append(link_elements[0].get_attribute("href"))

            # Ensure there are always 3 links
            top_three_links.extend([None] * (3 - len(top_three_links)))  # Pad with None if fewer than 3 links

            # If website was found, create a valid result row
            if website_rank is not None and link is not None:
                result_row = [website_to_check, filename, website_rank, link]
            else:
                result_row = [website_to_check, filename, 'не найдено', 'нет ссылки']

            result_row.extend(top_three_links)  # Add the top 3 links to the result
            results.append(result_row)

    # Convert to DataFrame and write to Excel
    df = pd.DataFrame(results,
                      columns=['Сайт', 'Запрос', 'Место в поиске', 'Ссылка', 'Ссылка1', 'Ссылка2', 'Ссылка3'])
    df.to_excel('yandex_novosibirsk.xlsx', index=False)


def test_find_competitors_and_website_rank_kazan(driver):
    results = []
    websites_to_check = ["101internet.ru"]
    for website_to_check in websites_to_check:
        for url, filename in zip(url_kazan, query_new_kazan):
            driver.get(url)
            search_results = driver.find_elements(By.CSS_SELECTOR, "li[data-cid]")
            top_three_links = []  # Store top three links
            website_rank = None
            link = None

            # Look for the website in the search results
            for result in search_results:
                link_elements = result.find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    url = link_elements[0].get_attribute("href")
                    if any(website in url for website in websites_to_check):
                        website_rank = search_results.index(result) + 1
                        link = url
                        break

            # Get the top three links
            for i in range(min(3, len(search_results))):
                link_elements = search_results[i].find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    top_three_links.append(link_elements[0].get_attribute("href"))

            # Ensure there are always 3 links
            top_three_links.extend([None] * (3 - len(top_three_links)))  # Pad with None if fewer than 3 links

            # If website was found, create a valid result row
            if website_rank is not None and link is not None:
                result_row = [website_to_check, filename, website_rank, link]
            else:
                result_row = [website_to_check, filename, 'не найдено', 'нет ссылки']

            result_row.extend(top_three_links)  # Add the top 3 links to the result
            results.append(result_row)

    # Convert to DataFrame and write to Excel
    df = pd.DataFrame(results,
                      columns=['Сайт', 'Запрос', 'Место в поиске', 'Ссылка', 'Ссылка1', 'Ссылка2', 'Ссылка3'])
    df.to_excel('yandex_kazan.xlsx', index=False)


def test_find_competitors_and_website_rank_samara(driver):
    results = []
    websites_to_check = ["101internet.ru"]
    for website_to_check in websites_to_check:
        for url, filename in zip(url_samara, query_new_samara):
            driver.get(url)
            search_results = driver.find_elements(By.CSS_SELECTOR, "li[data-cid]")
            top_three_links = []  # Store top three links
            website_rank = None
            link = None

            # Look for the website in the search results
            for result in search_results:
                link_elements = result.find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    url = link_elements[0].get_attribute("href")
                    if any(website in url for website in websites_to_check):
                        website_rank = search_results.index(result) + 1
                        link = url
                        break

            # Get the top three links
            for i in range(min(3, len(search_results))):
                link_elements = search_results[i].find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    top_three_links.append(link_elements[0].get_attribute("href"))

            # Ensure there are always 3 links
            top_three_links.extend([None] * (3 - len(top_three_links)))  # Pad with None if fewer than 3 links

            # If website was found, create a valid result row
            if website_rank is not None and link is not None:
                result_row = [website_to_check, filename, website_rank, link]
            else:
                result_row = [website_to_check, filename, 'не найдено', 'нет ссылки']

            result_row.extend(top_three_links)  # Add the top 3 links to the result
            results.append(result_row)

    # Convert to DataFrame and write to Excel
    df = pd.DataFrame(results,
                      columns=['Сайт', 'Запрос', 'Место в поиске', 'Ссылка', 'Ссылка1', 'Ссылка2', 'Ссылка3'])
    df.to_excel('yandex_samara.xlsx', index=False)