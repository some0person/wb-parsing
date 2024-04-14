import xlsxwriter

from time import sleep
from os import name as osName

from selenium import common
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait


OUTPUTFILE = "result.xlsx"
URL = r"https://www.wildberries.ru/catalog/elektronika/setevoe-oborudovanie"
MINGOODSCOUNT = 201  # Число необходимых товаров
RETRIES = 3  # Попытки нахождения элементов


def writer(goods: list[dict[str, str]]) -> None:
    workbook = xlsxwriter.Workbook(OUTPUTFILE)
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({"bold": True})
    worksheet.set_column(first_col=0, last_col=len(goods[0].keys()), width=40)

    for i, infoTitle in enumerate(goods[0].keys()):
        worksheet.write(0, i, infoTitle, bold)

    for e, element in enumerate(goods, start=1):
        for k, key in enumerate(element.keys()):
            worksheet.write(e, k, element[key])

    workbook.close()


def scrollPage(count: int, delay: float, driver: webdriver.firefox.webdriver.WebDriver) -> None:
    for i in range(1, count + 1):
        driver.execute_script(
            "window.scrollTo(0, document.body.scrollHeight)")
        sleep(delay)
        driver.execute_script(
            "window.scrollTo(0, 0)")
        sleep(delay)


def parsePages(driver: webdriver.firefox.webdriver.WebDriver) -> list[str]:
    driver.get(url=URL)
    sleep(0.5)
    driver.find_element(By.XPATH, "//button[contains(@class, 'cookies') and contains(@class, 'btn')]").click()

    parsedElements, completedRetries = [], 0
    while len(parsedElements) < MINGOODSCOUNT:
        if len(driver.find_elements(By.CLASS_NAME, "product-card")) == 0:
            break

        scrollPage(count=10, delay=0.25, driver=driver)
        pageElements = driver.find_elements(By.CLASS_NAME, "product-card")
        for element in pageElements:
            elementUrl = element.find_element(By.TAG_NAME, "a").get_attribute("href")
            parsedElements.append(elementUrl)

        try:
            driver.find_element(By.CLASS_NAME, "pagination-next").click()
        except common.NoSuchElementException:
            completedRetries += 1
            if completedRetries == RETRIES:
                break

    return parsedElements


def parseElements(parsedElements: list[str], driver: webdriver.firefox.webdriver.WebDriver) -> list[dict[str, str]]:
    elementsInfo = []
    
    for elementUrl in set(parsedElements):
        driver.get(url=elementUrl)
        info = {
            "Артикул": int(elementUrl.split("/")[-2]),
            "Изображение": driver.find_element(By.ID, "imageContainer").find_element(By.TAG_NAME, "img").get_attribute("src"),
            "Название": driver.find_element(By.TAG_NAME, "h1").text,
            "Продавец": driver.find_element(By.XPATH, "//a[contains(@class, 'seller-info') and contains(@class, 'name')]").text,
            "Цена": int(driver.find_element(By.XPATH, "//ins[contains(@class, 'price-block') and contains(@class, 'final-price')]").text[:-2].replace(" ", "")),
            "Оценка": driver.find_element(By.XPATH, "//span[contains(@class, 'product-review') and contains(@class, 'rating')]").text
        }
        
        driver.find_element(By.XPATH, "//button[contains(@class, 'product-page') and contains(@class, 'btn-detail')]").click()

        info["Описание"] = driver.find_element(By.XPATH, "//p[contains(@class, 'option') and contains(@class, 'text')]").text

        elementsInfo.append(info)
    
    return elementsInfo


def main() -> None:
    match osName:
        case "nt":
            driverPath = r"./drivers/geckodriver.exe"
        case _:
            exit("There is no driver for your system!")

    service = webdriver.FirefoxService(executable_path=driverPath)
    driver = webdriver.Firefox(service=service, keep_alive=True)
    driver.implicitly_wait(time_to_wait=10)  # Ожидание загрузки нужных элеметов

    parsedElements = parsePages(driver=driver)
    elementsInfo = parseElements(parsedElements=parsedElements, driver=driver)
    driver.quit()
    writer(elementsInfo)


if __name__ == "__main__":
    print(f"Parsing at least {MINGOODSCOUNT} elements... | {URL} to {OUTPUTFILE}")
    main()
    print("Done. Check result.xlsx file.")
