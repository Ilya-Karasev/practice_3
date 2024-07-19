import time
from pandas import read_excel
from tqdm import tqdm
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException

def connect_web(url_adress):
    """
    Создаем элемент класса webdriver и переходим на страницу по url-адресу
    :param url_adress: url-адрес сайта
    :return: driver: элемент класса webdriver
    """
    driver = webdriver.Chrome()
    driver.get(url_adress)
    time.sleep(5)
    return driver

def search_input(driver, delay, adress):
    """
    Функция для внесения данных в поисковую форму
    :param driver: элемент класса webdriver
    :param delay: время ожидания
    :param adress: адрес
    """
    try:
        # Поиск формы ввода на сайте
        elem_search_string = WebDriverWait(driver, delay).until(
            EC.presence_of_element_located((By.XPATH, "//input[@class='input__control _bold']"))
        )
        # Вписываем данные в форму
        elem_search_string.send_keys(adress)
        # Запускаем поиск
        elem_search_string.send_keys(Keys.ENTER)
    except Exception as ERROR_search_input:
        print(f'{adress} - не отработал. Ошибка: {ERROR_search_input}')

def clear_input_form(driver, delay):
    """
    Очистка формы ввода на сайте
    :param driver: элемент класса webdriver
    :param delay: время ожидания
    """
    attempt = 0
    while attempt < 3:
        try:
            elem_clear = WebDriverWait(driver, delay).until(
                EC.presence_of_element_located((By.XPATH, "//a[@class='small-search-form-view__pin']"))
            )
            elem_clear.click()
            break
        except StaleElementReferenceException:
            attempt += 1
        except TimeoutException:
            attempt += 1
        try:
            elem_clear = WebDriverWait(driver, delay).until(
                EC.presence_of_element_located((By.XPATH, "//div[@class='small-search-form-view__icon _type_close']"))
            )
            elem_clear.click()
            break
        except StaleElementReferenceException:
            attempt += 1
        except TimeoutException:
            attempt += 1

def work_selenium(np_search_adress, url_adress):
    """
    Основная функция работы скрипта
    :param np_search_adress: список адресов
    :param url_adress: url-адрес сайта
    """
    driver = connect_web(url_adress)  # Настройка selenium
    delay = 1  # Время ожидания

    for adress in tqdm(np_search_adress):
        time.sleep(1)
        search_input(driver, delay, adress)  # Вносим адрес в поисковую форму
        time.sleep(2)  # Ждем некоторое время для завершения поиска
        clear_input_form(driver, delay)  # Очищаем форму ввода для следующего поиска

    driver.quit()

if __name__ == '__main__':
    url_adress = 'https://yandex.ru/maps'  # Сайт "Яндекс карты"
    df_start_address = read_excel('Исправленные адреса.xlsx')  # Полученные адреса (после работы тетрадки)
    work_selenium(df_start_address['formating_adress'].dropna().tolist(), url_adress)  # Запуск основной функции