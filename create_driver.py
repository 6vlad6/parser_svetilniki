from selenium import webdriver


def create_driver(url):
    """
    Создание нового driver-объекта для парсинга новой страницы
    :param url: url-адрес
    :return: driver
    """
    options = webdriver.ChromeOptions()
    options.add_argument("start-maximized")
    # options.add_argument("--headless")  # работа без открытия браузера
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/105.0.0.0 Safari/537.36")
    options.add_argument("--disable-blink-features=AutomationControlled")

    driver = webdriver.Chrome(options=options)

    driver.get(url)

    return driver
