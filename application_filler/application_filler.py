#Импорт библиотек
from transliterate import translit
import win32com.client
#import random

class text_inst():
    #Функция поиска первого числа в строке
    def number_pos(self, string):
        for i, c in enumerate(string):
            if c.isdigit():
                return i
                break

#Выделение строчек из письма
strings = []
outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")
inbox = mapi.GetDefaultFolder(6)
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)
for message in messages:
    if (message.UnRead) and \
    ((message.subject.startswith("FW: В ответственность Вашей команды назначена заявка")) or \
    (message.subject.startswith("В ответственность Вашей команды назначена заявка"))) and \
    (message.body.find('ПРИЕМ НА РАБОТУ') > -1):
        string_pos_start = message.body.find('ФИО')
        string = message.body[string_pos_start::]
        string = string[:len(string) - 13]
        message.Unread = False
        strings.append(string)

#Формирование текстовых полей для заполнения
symbols = ['!', '#', '$', '%', '+', '&', '@']
for string in strings:
    string =  string[50:]
    date_start = text_inst().number_pos(string)
    fio = string[:date_start]
    surname, name, patronymic = fio.split()
    birth = string[len(string)-10::]

    #Формирование логина
    login_base = translit(surname, 'ru',  reversed=True).lower()
    login_init1 =  translit(name[0], 'ru',  reversed=True).lower()
    login_init2 =  translit(patronymic[0], 'ru',  reversed=True).lower()
    login = login_base + '-' + login_init1 + login_init2

    #Формирование пароля
    #random_num = random.randint(1000,9999)
    #random_symbol = random.choice(symbols)
    #password = login_base.capitalize() + str(random_num) + random_symbol
    password = 'Zaq123456&'