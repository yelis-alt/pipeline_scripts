#Импорт библиотек
from transliterate import translit
import win32com.client
import pyautogui
import pyperclip
import time
import random

class text_inst():
    #Метод поиска первого числа в строке
    def number_pos(self, string ):
        for i, c in enumerate(string):
            if c.isdigit():
                return i
                break
    #Метод вставки текста в поле
    def field_paste(self, field):
        pyperclip.copy(field)
        pyautogui.hotkey('ctrl', 'v')
text = text_inst()

class key_inst():
    #Метод повторения нажатия клавиш
    def cycle(self, count, key):
        for i in range(count):
            pyautogui.press(key)
key = key_inst()


#Выделение строчек из письма
strings = []
outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")
inbox = mapi.GetDefaultFolder(6)
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)
subject = 'В ответственность Вашей команды назначена заявка'
for message in messages:
    if (message.UnRead) and \
    ((message.subject.startswith('FW: ' + subject)) or \
    (message.subject.startswith(subject))) and \
    (message.body.find('ПРИЕМ НА РАБОТУ') > -1):
        string_pos_start = message.body.find('ФИО')
        string = message.body[string_pos_start::]
        string = string[:len(string) - 13]
        message.Unread = False
        strings.append(string)

#Формирование текстовых полей для заполнения
symbols = ['!', '#', '$', '%', '+', '&', '@']
for k, string in enumerate(strings):
    string = string[50:]
    date_start = text.number_pos(string)
    fio = string[:date_start]
    surname, name, patronymic = fio.split()

    #Формирование логина
    login_base = translit(surname, 'ru',  reversed=True).lower()
    login_init1 = translit(name[0], 'ru',  reversed=True).lower()
    login_init2 = translit(patronymic[0], 'ru',  reversed=True).lower()
    login = login_base + '-' + login_init1 + login_init2
    login = login.replace("'", '')

    #Формирование пароля
    random_num = random.randint(1000,9999)
    random_symbol = random.choice(symbols)
    password = login_base.capitalize() + str(random_num) + random_symbol

    #Активация окна активной директории
    window = pyautogui.getWindowsWithTitle("Active Directory - пользователи и компьютеры")
    window[0].activate()

    #Открытие формы
    if k == 0:
        key.cycle(10, 'left')
        pyautogui.press('right')
        key.cycle(2, 'down')
        pyautogui.press('right')
        key.cycle(17, 'down')
        pyautogui.press('right')
        key.cycle(9, 'down')
        pyautogui.hotkey('shift', 'f10')
        key.cycle(4, 'down')
        key.cycle(2, 'enter')
    else:
        pyautogui.hotkey('shift', 'f10')
        key.cycle(4, 'down')
        key.cycle(2, 'enter')

    #Заполнение формы
    time.sleep(3)
    text.field_paste(name)
    pyautogui.press('tab')
    text.field_paste(name[0])
    text.field_paste(patronymic[0])
    pyautogui.press('tab')
    text.field_paste(surname)
    pyautogui.press('tab')
    pyautogui.press('del')
    text.field_paste(surname)
    text.field_paste(' ')
    text.field_paste(name)
    text.field_paste(' ')
    text.field_paste(patronymic)
    pyautogui.press('tab')
    text.field_paste(login)
    key.cycle(4, 'tab')
    pyautogui.press('enter')
    text.field_paste(password)
    pyautogui.press('tab')
    text.field_paste(password)
    key.cycle(7, 'tab')
    key.cycle(6, 'enter')
    pyperclip.copy('')

    #Ожидание закрытия окна
    while True:
        try:
            pyautogui.getWindowsWithTitle("Новый объект - Пользователь")[0]
        except:
            break