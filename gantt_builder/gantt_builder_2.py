#Импорт библиотек
import plotly.express as px
import pandas as pd
import numpy as np

#Функция изъятия первых n слов
def three_words(df, col, n):
    new = df[col].str.split(' ')
    new = new.str[0:n]
    new = new.apply(lambda x: ' '.join(x))
    return new

#Подготовка переменных
path = './otchet.xlsx'
ind = []

#Изъятие элементов иерархии этапов и заданий
plan = pd.read_excel(path, sheet_name = 'Исход')
cols = ['Этап', 'Функция', 'Задача', 'Ответственный', 'Статус', 'Длительность', 'Начало', 'Окончание']
plan = plan[cols].replace(np.nan, '...')
plan['Статус'] = plan['Статус'].replace('...', 'Ждёт выполнения')
plan['Задача_'] = plan['Этап'].str[:7] + '/' + \
                  three_words(plan, 'Функция', 3) + '/' + \
                  three_words(plan, 'Задача', 3) + ' (' + \
                  plan['Длительность'] + ')'
plan = plan[['Задача_', 'Длительность', 'Начало', 'Окончание', 'Статус', 'Ответственный']]
plan = plan.rename({'Задача_': 'Задача'}, axis=1)
plan['Начало']= pd.to_datetime(plan['Начало'].replace('...', '10.02.2023'))
plan['Окончание']= pd.to_datetime(plan['Окончание'].replace('...', '11.02.2023'))

#Составление диаграммы
for i in range(plan.shape[0]):
    task = plan['Задача'][i]
    start = plan['Начало'][i]
    finish = plan['Окончание'][i]
    worker = plan['Ответственный'][i]
    status = plan['Статус'][i]
    row = dict(Задание = task,
               Начало = start,
               Конец = finish,
               Работник=worker,
               Статус=status)
    ind.append(row)

#Построение диаграммы Ганта
ind.sort(key=lambda item:(item['Начало'], item['Статус']), reverse = True)
job = pd.DataFrame(ind)
fig = px.timeline(job, x_start="Начало",
                       x_end="Конец",
                       y="Задание",
                       color="Работник",
                       pattern_shape="Статус",
                       pattern_shape_sequence=["", "x", "\\"],
                       height=3000)
fig.update_layout(
                  xaxis_tickformat = '%d.%m.%y',
                  yaxis_title=None,
                  xaxis = {'side': 'top'},
                 )
fig.update_xaxes(
                 dtick="M1"
                )

#Выгрузка диаграммы Ганта
fig.write_html("./gantt_chart_2.html")