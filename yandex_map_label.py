#Библиотеки
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium import webdriver
import pyperclip
import pyautogui
import time
import os
 
class web_rpa():
    #Метод ожидания появления элемента
    def waitForLoad(self, inputXPath):
        Wait = WebDriverWait(driver, 60)
        Wait.until(EC.element_to_be_clickable((By.XPATH, inputXPath)))
    #Метод проверки существования элемента
    def checkExistence(self, inputXPath):
        try:
            check = driver.find_element('xpath', inputXPath)
        except:
            return False
        return True
web = web_rpa()

#Формирование драйвера
options = webdriver.ChromeOptions()
codec = Service('./resources/yandexdriver.exe')
driver = webdriver.Chrome(service = codec, options=options)

#Подготовка к действиям
start = 83 ##начальная метка закраски
url = 'https://yandex.ru/map-constructor'
path = os.path.abspath("./resources/dots.xlsx").replace('\x5c', '//')
driver.get(url)
time.sleep(2)

#Открытие конструктора
opens = driver.find_element('xpath', '//span[@class="button__text" and text()="Открыть карту"]')
opens.click()
waited_element = '//span[@class="button__text" and text()="К импорту →"]'
web.waitForLoad(waited_element)

#Загрузка файла с координатами
xpaths = [waited_element,
          '//input[@title="Выбрать файл"]']
for i in range(len(xpaths)):
    element = driver.find_element('xpath', xpaths[i])
    element.click()
pyperclip.copy(path)
pyperclip.paste()
pyautogui.press('enter')
time.sleep(3)

#Перекрашивание добавленных точек
ids = f'//div[contains(@data-id, "_{start}")]'
while web.checkExistence(ids):
    xpaths = [ids,
              '(//span[@class="icon select__tick"])[4]',
              '//div[@class="menu__item menu__item_theme_colors menu__item_index_7 i-bem menu__item_js_inited"]']
    for i in range(len(xpaths)):
        element = driver.find_element('xpath', xpaths[i])
        element.click()
    start += 1
    ids = f'//div[contains(@data-id, "_{start}")]'

#Сохранение результата:
xpaths = ['//span[@class="button__text" and text()="Сохранить и продолжить"]',
          '//a[text()="Экспорт"]',
          '//span[@class="button__text" and text()="KML"]',
          '//span[@class="button__text" and text()="Скачать"]']
for i in range(len(xpaths)):
    time.sleep(2)
    element = driver.find_element('xpath', xpaths[i])
    element.click()
time.sleep(5)
driver.close()