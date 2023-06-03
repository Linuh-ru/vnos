# selenium 4
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromiumService
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.core.utils import ChromeType
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from tkinter import filedialog as fd


#Извлечение документа из файла word
from docx import Document
doc = Document(fd.askopenfilename())


#последовательность всех таблиц документа
all_tables = doc.tables
#print('Всего таблиц в документе:', len(all_tables))


#функция
def upload(all_tables, ticket):

    #загрузка браузера
    options = Options()
    options.add_experimental_option("detach", True)
    driver = webdriver.Chrome(service=ChromiumService(ChromeDriverManager(chrome_type=ChromeType.CHROMIUM).install()),
                              options=options)
    driver.implicitly_wait(0.5)

    #Стартовая страница (аутентификация)
    driver.get("http://tasks.o2dc.ru/login.php")
    username = driver.find_element(by=By.NAME, value="username")
    username.send_keys("aomarov")
    password = driver.find_element(by=By.NAME, value="password")
    password.send_keys("qQ12345")
    password.send_keys(Keys.ENTER)

    # зайти в тикет
    # driver.implicitly_wait(10)
    time.sleep(5)
    driver.get(f"http://tasks.o2dc.ru/ticket/{ticket}/")

    # Нажать на кнопку внос
    driver.implicitly_wait(10)
    actual = driver.find_element(By.CSS_SELECTOR, 'a[data-target="#ModalCompanyIN"]')
    # driver.execute_script("arguments[0].scrollIntoView();", actual)
    driver.execute_script("arguments[0].click();", actual)

    #выгрузка таблицы (ячейки с 1-ой по 5-ую)
    n = 1
    for row in all_tables[1].rows[1::]:
        list = []
        for cell in row.cells[1::]:
            list.append(cell.text)
        print(list)

        #Заполнение таблицы
        model = driver.find_element(By.XPATH, f"//table/tbody/tr[@data-device='{n}']/td/input[@data-content='device_name']").send_keys(list[0])
        sn = driver.find_element(By.XPATH, f"//table/tbody/tr[@data-device='{n}']/td/input[@data-content='device_snum']").send_keys(list[1])
        count = driver.find_element(By.XPATH, f"//table/tbody/tr[@data-device='{n}']/td/input[@data-content='device_ucount']").send_keys(list[3])
        unit = driver.find_element(By.XPATH, f"//table/tbody/tr[@data-device='{n}']/td/input[@data-content='device_count']").send_keys(list[2])
        n += 1
        actual = driver.find_element(By.CSS_SELECTOR, 'a[data-type="ClientDeviceNewRowSelector"]')
        driver.execute_script("arguments[0].click();", actual)


#проверка соответствия таблице
def check(all_tables):
    for row in all_tables[1].rows[1::]:
        list = []
        for cell in row.cells[1::]:
            list.append(cell.text)
        print(list)
check(all_tables)

#upload(all_tables, 125597)