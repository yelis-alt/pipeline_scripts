#Библиотеки
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium import webdriver
import pandas as pd
import requests
import datetime
import time
import re

#Логирование статуса вебсайта
def logize(num):
    with open("./log.txt",'w') as f:
        pass
        f.write(num)

#Функция ожидания загрузки страницы 15 секунд
def waitForLoad(inputXPath):
    Wait = WebDriverWait(driver, 15)
    Wait.until(EC.element_to_be_clickable((By.XPATH, inputXPath)))

#Функция записи недоступных элементов в массив
def nulled(N):
    for i in range(N):
        history.append(0)

#Функция проверки доступности элементов
def check_elements():
    driver.get(url)
    #Вход в личный кабинет
    private_office = driver.find_element('xpath', '//span[@class="ui-button-text ui-c" and contains(text(), "Личный кабинет")]')
    try:
        private_office.click()
        history.append(1)
    except:
        history.append(0)
    if history[0] == 1: #Проверка доступности логина
        #Авторизация
        try:
            waitForLoad('//input[@id="workplaceTopForm:j_mail_login"]')
            login_field = driver.find_element('xpath', '//input[@id="workplaceTopForm:j_mail_login"]')
            login_field.send_keys('211714@edu.fa.ru')
            password_field = driver.find_element('xpath', '//input[@id="workplaceTopForm:j_password"]')
            password_field.send_keys('Test1234!')
            login_button = driver.find_element('xpath', '//button[@id="workplaceTopForm:loginBtn"]')
            login_button.click()
            history.append(1)
        except:
            history.append(0)
        if history[1] == 1: #Проверка успешности авторизации
            #Элементы личного кабинета
            try:
                waitForLoad('//button[@id="workplaceTopForm:buttonLK"]')
                name_field = driver.find_element('xpath', '//button[@id="workplaceTopForm:buttonLK"]')
                name_field.click()
                private_office = driver.find_element('xpath', '//a[@id="workplaceTopForm:j_idt661"]')
                private_office.click()
                waitForLoad('//a[contains(text(), "Заявки и обращения")]')
                po_elements = ['//a[contains(text(), "Заявки и обращения")]',
                               '//a[contains(text(), "Опросы")]',
                               '//a[contains(text(), "Договоры")]',
                               '//a[contains(text(), "Уведомления")]',
                               '//a[contains(text(), "История действий")]',
                               '//a[contains(text(), "Подписка")]',
                               '//div[@class="metering_tabs"]']
                po_waits = ['//input[@id="workplaceForm:messagesProfile:myRARProfile:fromDate_input"]',
                            '//a[contains(text(), "Результаты опросов")]',
                            '(//li[@value="0"])[2]',
                            '//input[@id="workplaceForm:messagesProfile:selectedFromUsr_input"]',
                            '//input[@id="workplaceForm:messagesProfile:selectedFromAudit_input"]',
                            '//a[@id="workplaceForm:messagesProfile:addSuscribeButton"]',
                            '//a[@class="ui-link ui-widget link-as-btn"]']
                for i in range(len(po_elements)):
                    try:
                        private_office_el = driver.find_element('xpath', po_elements[i])
                        private_office_el.click()
                        waitForLoad(po_waits[i])
                        time.sleep(2)
                        history.append(1)
                    except:
                        history.append(0)
            except:
                nulled(num_el-5)
            #Калькуляторы:
            calculators = ['//span[contains(text(), "Калькулятор необходимой мощности")]',
                           '//span[contains(text(), "Калькулятор стоимости установки приборов учёта")]',
                           '//span[contains(text(), "Калькулятор стоимости дополнительных услуг")]']
            calculator_checks = ['//div[@class="ui-radiobutton-box ui-widget ui-corner-all ui-state-default"]',
                                 '//label[@id="workplaceForm:calc_regions_label"]',
                                 '//label[@id="workplaceForm:j_idt11178_label"]']
            for i in range(len(calculators)):
                try:
                    cost_field = driver.find_element('xpath', '//span[@class="ui-menuitem-text" and contains(text(), "Стоимость")]')
                    cost_field.click()
                    calculator_type = driver.find_element('xpath', calculators[i])
                    calculator_type.click()
                    waitForLoad(calculator_checks[i])
                    history.append(1)
                    time.sleep(2)
                except:
                    history.append(0)
        else:
            nulled(num_el-2)
    else:
        nulled(num_el-1)

#Функция формирвоания сообщения бота
def bot_message(status, history):
    now = datetime.datetime.now()
    day = now.date().strftime("%d/%m/%Y").replace('/', '.')
    hour = now.strftime("%H:%M:%S")
    add = '\n' + '________________________________'
    main_text = f'Уведомляю Вас, что по состоянию на {day} в {hour} доступ к порталу https://портал-тп.рф'

    if status == 'down':
        text = main_text + ' прекращён' + add
    elif status == 'up':
        text = main_text + ' возобновлён' + add
    elif status == 'still_down':
        text = main_text + ' отсутствует' + add
    else:
        text = main_text + ' функционирует' + add

    elements_array = ['>Форма для ввода логина и пароля: ',
                      '>Авторизация: ',
                     #'>Элементы личного кабинета:',
                          ' -Заявки и обращения: ',
                          ' -Опросы: ',
                          ' -Договоры: ',
                          ' -Уведомления: ',
                          ' -История действий: ',
                          ' -Подписка: ',
                          ' -Потребление э/э: ',
                      '>Калькулятор необходимой мощности: ',
                      '>Калькулятор стоимости установки приборов учёта: ',
                      '>Калькулятор стоимости дополнительных услуг: ']

    for i in range(len(history)):
        if i == 1:
            text = text + '\n' + elements_array[i] +  str(history[i]) + ';' + '\n' + '>Элементы личного кабинета:'
        else:
            text = text + '\n' + elements_array[i] + str(history[i]) + ';'
            text = text.replace(': 1', ': доступно').replace(': 0', ': не доступно')
    print(text)

#Функция повторения проверки
def check_elements_again():
    history = []
    time.sleep(300)
    check_elements()

#Считывание статуса предыдущей итерации
log = int(pd.read_csv('./log.txt').columns[0])
options = webdriver.ChromeOptions()
#options.add_argument('headless')
codec = Service('./yandexdriver.exe')
driver = webdriver.Chrome(service = codec, options=options)

#Подготовка переменных
num_el = 12
url = 'https://портал-тп.рф/platform/portal/tehprisEE_portal'
tel = 'https://api.telegram.org/https://t.me/RD_Zabbix_bot/sendMessage?chat_id=-621157789'
history = []

#Извлечение
status = requests.head(url)
code = int(re.search(r"\[([A-Za-z0-9_]+)\]",
       str(status))[1])

#Проверка доступности страницы и её элементов
for i in range(1):
    if code == 200:
        check_elements()
        if (sum(history[0:2]) == 2) & (sum(history[2::]) == 0): #на случай моментного сбоя
            try:
                check_elements_again() #задержка 5 минут
            except:
                bot_message('down', history)
                logize('0')
                break
        if log == 0:
            bot_message('up', history)
            logize('1')
        else:
            bot_message('still_up', history)
    else:
        if log == 1:
            bot_message('down', history)
            logize('0')
        else:
            bot_message('still_down', history)