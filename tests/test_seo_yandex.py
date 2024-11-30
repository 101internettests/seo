from selenium.webdriver.common.by import By
import time
from pages.data_google import urlsyandex, path_yandex
from pages.data_yandex import url_ms, path_yandex_msk, url_spb, path_yandex_spb, url_ufa, path_yandex_ufa
from pages.data_yandex_second import url_krasnodar, path_yandex_krasnodar, url_novosibirsk,path_yandex_novosibirsk
import pandas as pd


def test_find_competitors_and_website_rank_msk(driver):
    results = []
    websites_to_check = ["moskvaonline.ru"]
    for website_to_check in websites_to_check:
        for url, filename in zip(url_ms, path_yandex_msk):
            driver.get(url)
            search_results = driver.find_elements(By.CSS_SELECTOR, "li[data-cid]")
            top_three_links = []  # Store top three links
            website_rank = None
            link = None

            for result in search_results:
                link_elements = result.find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    url = link_elements[0].get_attribute("href")
                    if any(website in url for website in websites_to_check):
                        website_rank = search_results.index(result) + 1
                        link = url
                        break

            for i in range(min(3, len(search_results))):
                link_elements = search_results[i].find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    top_three_links.append(link_elements[0].get_attribute("href"))

            if website_rank is not None and link is not None:
                result_row = [website_to_check, filename, website_rank, link]
                result_row.extend(top_three_links)
                results.append(result_row)
            else:
                result_row = [website_to_check, filename, 'не найдено', 'нет ссылки']
                result_row.extend(top_three_links)
                results.append(result_row)

            df = pd.DataFrame(results,
                              columns=['Сайт', 'Запрос', 'Место в поиске', 'Ссылка', 'Ссылка1', 'Ссылка2', 'Ссылка3'])
            df.to_excel('yandex_msk.xlsx', index=False)
            time.sleep(2)


def test_find_competitors_and_website_rank_spb(driver):
    results = []
    websites_to_check = ["piter-online.net"]
    for website_to_check in websites_to_check:
        for url, filename in zip(url_spb, path_yandex_spb):
            driver.get(url)
            search_results = driver.find_elements(By.CSS_SELECTOR, "li[data-cid]")
            top_three_links = []  # Store top three links
            website_rank = None
            link = None

            for result in search_results:
                link_elements = result.find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    url = link_elements[0].get_attribute("href")
                    if any(website in url for website in websites_to_check):
                        website_rank = search_results.index(result) + 1
                        link = url
                        break

            for i in range(min(3, len(search_results))):
                link_elements = search_results[i].find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    top_three_links.append(link_elements[0].get_attribute("href"))

            if website_rank is not None and link is not None:
                result_row = [website_to_check, filename, website_rank, link]
                result_row.extend(top_three_links)
                results.append(result_row)
            else:
                result_row = [website_to_check, filename, 'не найдено', 'нет ссылки']
                result_row.extend(top_three_links)
                results.append(result_row)

            df = pd.DataFrame(results,
                              columns=['Сайт', 'Запрос', 'Место в поиске', 'Ссылка', 'Ссылка1', 'Ссылка2', 'Ссылка3'])
            df.to_excel('yandex_spb.xlsx', index=False)
            time.sleep(2)


def test_find_competitors_and_website_rank_ufa(driver):
    results = []
    websites_to_check = ["101internet.ru"]
    for website_to_check in websites_to_check:
        for url, filename in zip(url_ufa, path_yandex_ufa):
            driver.get(url)
            search_results = driver.find_elements(By.CSS_SELECTOR, "li[data-cid]")
            top_three_links = []  # Store top three links
            website_rank = None
            link = None

            for result in search_results:
                link_elements = result.find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    url = link_elements[0].get_attribute("href")
                    if any(website in url for website in websites_to_check):
                        website_rank = search_results.index(result) + 1
                        link = url
                        break

            for i in range(min(3, len(search_results))):
                link_elements = search_results[i].find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    top_three_links.append(link_elements[0].get_attribute("href"))

            if website_rank is not None and link is not None:
                result_row = [website_to_check, filename, website_rank, link]
                result_row.extend(top_three_links)
                results.append(result_row)
            else:
                result_row = [website_to_check, filename, 'не найдено', 'нет ссылки']
                result_row.extend(top_three_links)
                results.append(result_row)

            df = pd.DataFrame(results,
                              columns=['Сайт', 'Запрос', 'Место в поиске', 'Ссылка', 'Ссылка1', 'Ссылка2', 'Ссылка3'])
            df.to_excel('yandex_ufa.xlsx', index=False)
            time.sleep(2)


def test_find_competitors_and_website_rank_krasnodar(driver):
    results = []
    websites_to_check = ["101internet.ru"]
    for website_to_check in websites_to_check:
        for url, filename in zip(url_krasnodar, path_yandex_krasnodar):
            driver.get(url)
            search_results = driver.find_elements(By.CSS_SELECTOR, "li[data-cid]")
            top_three_links = []  # Store top three links
            website_rank = None
            link = None

            for result in search_results:
                link_elements = result.find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    url = link_elements[0].get_attribute("href")
                    if any(website in url for website in websites_to_check):
                        website_rank = search_results.index(result) + 1
                        link = url
                        break

            for i in range(min(3, len(search_results))):
                link_elements = search_results[i].find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    top_three_links.append(link_elements[0].get_attribute("href"))

            if website_rank is not None and link is not None:
                result_row = [website_to_check, filename, website_rank, link]
                result_row.extend(top_three_links)
                results.append(result_row)
            else:
                result_row = [website_to_check, filename, 'не найдено', 'нет ссылки']
                result_row.extend(top_three_links)
                results.append(result_row)

            df = pd.DataFrame(results,
                              columns=['Сайт', 'Запрос', 'Место в поиске', 'Ссылка', 'Ссылка1', 'Ссылка2', 'Ссылка3'])
            df.to_excel('yandex_krasnodar.xlsx', index=False)
            time.sleep(2)


def test_find_competitors_and_website_rank_novosibirsk(driver):
    results = []
    websites_to_check = ["101internet.ru"]
    for website_to_check in websites_to_check:
        for url, filename in zip(url_novosibirsk, path_yandex_novosibirsk):
            driver.get(url)
            search_results = driver.find_elements(By.CSS_SELECTOR, "li[data-cid]")
            top_three_links = []  # Store top three links
            website_rank = None
            link = None

            for result in search_results:
                link_elements = result.find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    url = link_elements[0].get_attribute("href")
                    if any(website in url for website in websites_to_check):
                        website_rank = search_results.index(result) + 1
                        link = url
                        break

            for i in range(min(3, len(search_results))):
                link_elements = search_results[i].find_elements(By.CSS_SELECTOR, "a")
                if link_elements:
                    top_three_links.append(link_elements[0].get_attribute("href"))

            if website_rank is not None and link is not None:
                result_row = [website_to_check, filename, website_rank, link]
                result_row.extend(top_three_links)
                results.append(result_row)
            else:
                result_row = [website_to_check, filename, 'не найдено', 'нет ссылки']
                result_row.extend(top_three_links)
                results.append(result_row)

            df = pd.DataFrame(results,
                              columns=['Сайт', 'Запрос', 'Место в поиске', 'Ссылка', 'Ссылка1', 'Ссылка2', 'Ссылка3'])
            df.to_excel('yandex_novosibirsk.xlsx', index=False)
            time.sleep(2)