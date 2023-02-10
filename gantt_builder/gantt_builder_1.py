#Импорт библиотек
from datetime import timedelta
import plotly.express as px
import pandas as pd

#Подготовка переменных
works = []
ind = []
path = './plan.xlsx'

#Изъятие элементов иерархии этапов и заданий
plan = pd.read_excel(path)
names = ['Этапы договора', 'Название']
types = plan.columns[plan.shape[1]-16: plan.shape[1]].to_list()
cols = names + types
plan = plan[cols]

plan = plan.melt(id_vars=names,
                 var_name="Тип",
                 value_name="Дата")

stages = list(plan['Этапы договора'].unique())
parts = list(plan['Название'].unique())
start_txt = 'Дата начала '
finish_txt = 'Дата окончания '
works_ = [x.replace(start_txt, '').replace(finish_txt, '') for x in types]
for i in range(len(works_)):
    if i % 2 !=0:
        works.append(works_[i])
plan_copy = plan.copy()

#Формирование списка работ со сроками
for stage in stages:
    for part in parts:
        temp = plan_copy
        temp = temp.loc[(temp['Этапы договора'] == stage) & (temp['Название'] == part)].reset_index(drop=True)
        if temp.shape[0]>0:
            for work in works:
                start = temp.loc[temp['Тип'] == start_txt+work].iloc[0,temp.shape[1]-1]
                finish = temp.loc[temp['Тип'] == finish_txt+work].iloc[0,temp.shape[1]-1]
                if start == finish:
                    finish += timedelta(days=1)
                task = " ".join(stage.split(' ')[0:2]) + '/' + \
                       part.split(' ')[0] + '/' + \
                       work
                row = dict(Задание = task,
                           Начало = start,
                           Конец = finish)
                ind.append(row)

#Построение диаграммы Ганта
ind.sort(key=lambda item:item['Начало'], reverse = True)
job = pd.DataFrame(ind)
fig = px.timeline(job, x_start="Начало",
                       x_end="Конец",
                       y="Задание",
                       height = 5000)
fig.update_layout(
                  xaxis_tickformat = '%d.%m.%y',
                  yaxis_title=None,
                  xaxis = {'side': 'top'},
                 )
fig.update_xaxes(
                 dtick="M1"
                )

#Выгрузка диаграммы Ганта
fig.write_html("./gantt_chart_1.html")