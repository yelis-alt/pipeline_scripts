#Библиотеки
from datetime import datetime
import pandas as pd
import numpy as np
import calendar
import warnings
import json
import os
warnings.filterwarnings('ignore')

#Метод изъятия диапазона продолжительности
def date_rangize(start, finish):
    period_start = pd.Series(pd.to_datetime(start.replace(' ', ''),
                                            dayfirst=True)).min()
    period_s = period_start.replace(day=1)
    period_finish = pd.Series(pd.to_datetime(finish.replace(' ', ''),
                                             dayfirst=True)).max()
    period_f = period_finish.replace(day=2)
    if (pd.isnull(period_start) == False) & (pd.isnull(period_finish) == False):
        period = pd.date_range(period_s,
                               period_f,
                               freq='MS').strftime("%Y-%b").tolist()
    else:
        period = 0
    res = [period_start, period_finish, period]
    return res

#Метод индексации столбцов
def col_indexize(tab, action, add):
    if action == 'indexing':
        col_index = [name + f'_{lvl + add}' for name in list(tab.columns)]
        tab.columns = col_index
    elif action == 'deindexing':
        col_index = [name.replace(f'_{lvl + add}', '') for name in list(tab.columns)]
        col_index = [x if x!=task else f'{task}_{lvl + add}' for x in col_index]
        tab.columns = col_index
    return tab

#Класс методов по формированию html-таблицы
class html_formatter():
    def tr_expand(self, grandparent, parent, expand,
                        ids, style, collapse,
                        title, n_items, row):
        row_append = f'<tr class="treegrid-{grandparent}{parent}{expand}" ' \
                     f'id="{ids}" {style}>' + '\n' \
                     f'<td style="padding: 2; border: 2px solid; ' \
                                f'text-align: left;">' + '\n' \
                     f'<span {collapse} </span>' + '\n' \
                     f'{title} ' + '\n' \
                     f'</td>'
        color = ''
        for key in color_dict.keys():
            if key in row:
                color = color_dict[key]
                break
        td = ''
        n_num = 0
        border_last = ''
        for n in range(n_items):
            if (isinstance(row[n], float)) & (row[n]!=float(0)):
                n_num += 1
            if n+1 == n_items:
                border_last= ' border-right: 2px solid;'
            if html_formatter.same_check(row): #пустые строки
                td += f'<td style="padding: 0; border-left: 3px solid {back};border-bottom: 2px solid black;{border_last}">' + '\n'
            elif row[n] == 100:
                css_class = 'class = "full"'
                td += f'<td {css_class} style="padding: 0; border: 1px solid; ' \
                      f'background-color: {color};{border_last}">' + '\n'
            elif row[n] == float(0):
                css_class = 'class = "null"'
                td += f'<td {css_class} style="padding: 0; border: 1px solid; ' \
                      f'background-color: {back};{border_last}">' + '\n'
            elif '|' in str(row[n]):
                begin = row[n].split('|')[0]
                end = row[n].split('|')[1]
                css_class = 'class = "inter"'
                td += f'<td {css_class} style="padding: 0; border: 1px solid;'\
                      f' background: linear-gradient(to right, {back} 0%, {back} {begin}%, ' \
                      f'{color} {begin}%, {color} {end}%, {back} {end}%, {back} 100%);{border_last}">' + '\n'
            elif (isinstance(row[n], float)) & (n_num==1):
                css_class = 'class = "start"'
                td += f'<td {css_class} style="padding: 0; border: 1px solid;'\
                      f' background: linear-gradient(to right, {back} 0%, {back} {100-row[n]}%, ' \
                      f'{color} {100-row[n]}%, {color} 100%);{border_last}">' + '\n'
            elif (isinstance(row[n], float)) & (n_num!=1):
                css_class = 'class = "finish"'
                td += f'<td {css_class} style="padding: 0; border: 1px solid;'\
                      f' background: linear-gradient(to right, {color} 0%, {color} {row[n]}%, ' \
                      f'{back} {row[n]}%, {back} 100%);{border_last}">' + '\n'
            else: #обычная ячейка таблицы
                td += f'<td style="padding: 1px; border: 2px solid">{row[n]}' + '\n'
            td += f'</td>'
        row_append += td + f'</tr>'
        return row_append

    def tr_build(self, all_cols):

         tr = '<thead>' + '\n'\
              '<tr id="tree-head-1">'
         for k, tag in enumerate(all_cols):
            row_width = ''
            if k == 0:
                row_width = ' style="width: 380px;"'
            tr += f'<th{row_width}>' + tag + '</th>'
         tr += '</tr>' + '\n'\
               '</thead>'
         return tr

    def same_check(self, list):
        return all(x == list[0] for x in list)
html_formatter = html_formatter()

#Входные переменные
aggr = 'ITEMS' #название агрегирующего поля
task = 'TITLE' #поле с названием процесса
n_levels = 5 #число уровней агрегации
select = {'STATUS': 'Статус', #поля для отбора с переводом
          'DURATION': 'Длительность',
          'BEGIN_DATE_FORMAT': 'Начало',
          'CLOSE_DATE_FORMAT': 'Окончание',
          'RESPONSOBLE_FIO': 'Ответственный'}

#Открытие json
with open('./resources/data.json') as fp:
    data = json.load(fp)

lvl = 0
fields_original = [task] + list(select.keys()) + [aggr]
tab = pd.json_normalize(data,max_level=1)[fields_original]
tab = col_indexize(tab, 'indexing', 0)
fields = list(tab.columns) + fields_original
tab_aggr = tab.copy()
tab_aggr = col_indexize(tab_aggr, 'deindexing', 0)

#Изъятие остальных уровней
for lvl in range(n_levels):
    tab = tab.explode(f'{aggr}_{lvl}', ignore_index=True)
    tab = pd.concat([tab, tab[f'{aggr}_{lvl}'].apply(pd.Series)], axis=1)
    if n_levels - lvl == 1:
        fields.remove(aggr)
        fields_original.remove(aggr)
    tab = tab[fields]
    col_index = [name + f'_{lvl+1}' for name in fields_original]
    cols = fields[:-len(fields_original)] + col_index
    tab.columns = cols
    fields = col_index + fields
    col_lvl = [x for x in list(tab.columns) if (int(x[-1])==lvl+1) | (x[:-2]==task)]
    tab_add = tab[col_lvl]
    tab_add = col_indexize(tab_add, 'deindexing', 1)
    tab_aggr[f'{task}_{lvl+1}'] = np.nan
    tab_aggr = pd.concat([tab_aggr, tab_add], axis = 0)

#Упорядочивание столбцов
col_aggr = [x for x in list(tab_aggr.columns) if x==aggr] #удаление агрегирующего столбца
tab_aggr = tab_aggr.drop(col_aggr, axis=1)
col_task = [x for x in list(tab_aggr.columns) if x[:-2]==task] #расположение задач в порядке возрастания уровня
col_task.sort()
col_calc = list(select.keys())
tab_aggr = tab_aggr[col_task + col_calc].drop_duplicates()
tab_aggr = tab_aggr.replace(np.nan, ' ')

#Сортировка агрегаций
sort_arr = ['Задачи договора',
                'Задачи этапа',
                'Функции этапа',
                'Документооборот этапа']
sort_replace = [str(k)*2 + x for k,x in enumerate(sort_arr)]
for k, repl in enumerate(sort_arr):
    tab_aggr = tab_aggr.replace(repl, sort_replace[k])

#Составление сводной таблицы
pivot = tab_aggr.pivot_table(index = col_task,
                             values = col_calc,
                             aggfunc ='sum')

#Сброс индекса
pivot.to_excel('./data.xlsx')
pivot = pd.read_excel('./data.xlsx').replace(np.nan, ' ')
os.remove('./data.xlsx')
pivot.rename(columns=select, inplace=True)
col_pivot = list(pivot.columns)
col_rename = [x.replace(f'{task}_', '') for x in col_pivot]
pivot.columns = col_rename
col_pivot = list(pivot.columns)

for k, repl in enumerate(sort_replace): #возвращение исходных названий столбцам
    pivot = pivot.replace(repl, sort_arr[k])

#Добавление отступов
expands = col_pivot[:(n_levels+1)]
for k, col in enumerate(expands):
    pivot[col] = np.where(pivot[col] == ' ',
                                        ' ',
                                        '__'*k + pivot[col].astype(str))
cols_prefinal = list(pivot.columns)

#Добавление временных столбцов
pivot_temp = pivot.copy()[pivot['0']!=' ']
start = pivot_temp['Начало']
finish = pivot_temp['Окончание']
month = date_rangize(start, finish)[2]

all_cols = pivot.columns.to_list() + month #формирование массива столбцов параметров
all_cols = ['Иерархия'] + [x for x in all_cols if str(x).isnumeric() == False]

#Заполнение временных столбцов
width = 50 #ширина шкалы прогресса
height = 25 #длина шкалы прогресса
back = 'azure' #фон диаграммы
font_size = 2.5 #размер шрифта
n_rows = pivot.shape[0]
col_begin = pivot.columns.get_loc('Начало')
col_end = pivot.columns.get_loc('Окончание')
for col in month:
    pivot[col] = float(0.00)
for col in month: #заполнение столбцов с месяцами величиной ширины
    col_month = pivot.columns.get_loc(col)
    for row in range(n_rows):
        start = pivot.iloc[row, col_begin]
        finish = pivot.iloc[row, col_end]
        range_attr = date_rangize(start, finish)
        period_start = range_attr[0]
        period_finish = range_attr[1]
        period = range_attr[2]
        if period != 0:
            period_nons = list(set(month) - set(period))
            if col not in period_nons:
                month_number = datetime.strptime(col[-3:], '%b').month
                year_number = int(col[:4])
                if (period_start.month == period_finish.month) & (period_start.year == period_finish.year):
                    width_val_start = round(period_start.day/month_range*100)
                    width_val_finish = round(period_finish.day/month_range*100)
                    pivot.iloc[row, col_month] = f'{width_val_start}|{width_val_finish}'
                elif (month_number == period_start.month) & (year_number == period_start.year): #проверка границы начала
                    month_range = calendar.monthrange(year_number, month_number)[1]
                    width_val = round((month_range-period_start.day)/month_range*100, 2)
                    pivot.iloc[row, col_month] = width_val
                elif (month_number == period_finish.month) & (year_number == period_finish.year): #проверка границы конца
                    month_range = calendar.monthrange(year_number, month_number)[1]
                    width_val = round(period_finish.day/month_range*100, 2)
                    pivot.iloc[row, col_month] = width_val
                else:
                    pivot.iloc[row, col_month] = 100
month_dict = {'Jan': 'Янв', 'Feb': 'Фев', 'Mar': 'Мар', #приведение даты к необходимому формату
              'Apr': 'Апр', 'May': 'Май', 'Jun': 'Июн',
              'Jul': 'Июл', 'Aug': 'Авг', 'Sep': 'Сен',
              'Oct': 'Окт', 'Nov': 'Ноя', 'Dec': 'Дек'}
for key, value in month_dict.items(): #перевод дат на русский
    month = [x.replace(key,value) for x in month]
month = [x[-3:] + '-' + x[:4] for x in month]
all_cols = all_cols[:-len(month)] + month

color_dict = {                         #задание цвета
               'Не начат': 'darkgrey',
               'Работы не начаты': 'darkgrey',
               'В работе': 'blue',
               'Выполняется': 'blue',
               'Новая (не просмотрена) ': 'purple',
               'Новая (не просмотрена)': 'purple',
               'Просрочена ': 'darkred',
               'Просрочена': 'darkred',
               'Почти просрочена ': 'red',
               'Почти просрочена': 'red',
               'Завершена ': 'green',
               'Завершена': 'green',
              }
#html-конвертация
pos = list(pivot.columns).index(all_cols[1]) - 1 #число уровней
n_items = len(all_cols)-1 #число заполняемых полей
header = '<!doctype html>' + '\n' \
         '<html>' + '\n' \
         '<head>' + '\n' \
         '<meta charset="utf-8">' + '\n' \
         '<title>Расписание</title>' + '\n' \
         '<style>' + '\n' \
         'table {border-collapse: collapse;'+ '\n' \
                'width: 100%; table-layout: fixed;}'+ '\n' \
         'thead tr:first-child th {'+ f'position: sticky; top: 0; background-color: {back};' + '}' + '\n' \
         'td {text-align: center; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;}'+ '\n' \
         'th {border: solid 2px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;}' + '\n' \
              '.treegrid-expander {width:10px; height: 10px; display: inline-block;' + '\n' \
                                   'position: relative;}' + '\n' \
              '.treegrid-expanded ' + '\n' \
              '.treegrid-expander {background-color: green;cursor: pointer;}' + '\n' \
              '.treegrid-collapsed ' + '\n' \
              '.treegrid-expander {background-color: red;cursor: pointer;}' + '\n' \
         '</style>' + '\n' \
         '</head>' + '\n' \
         f'<body style="background-color: {back};">' + '\n' \
         f'<font size="{font_size}" face="Helvetica">' + '\n' \
         '<table>' + '\n' \
         '<tbody>' + '\n' \
         '<tr>' + '\n' \
         '<td>' + '\n' \
         '<table id="tree-1" border="0">' + '\n' \
         '<tbody>' +\
         html_formatter.tr_build(all_cols)
footer = '</tbody>' + '\n' \
         '</table>' + '\n' \
         '</font>' + '\n' \
         '</td>' + '\n' \
         '</tr>' + '\n' \
         '</tbody>' + '\n' \
         '</table>' + '\n' \
         '<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.10.2/jquery.min.js"></script>' + '\n' \
         '<script>' + '\n' \
             '!function(a){var b={initTree:function(b){var c=a.extend({},this.treegrid.defaults,b);' + '\n' \
             'return this.each(function(){var b=a(this);b.treegrid("setTreeContainer",a(this)),b.treegrid("setSettings",c),' + '\n' \
             'c.getRootNodes.apply(this,[a(this)]).treegrid("initNode",c),b.treegrid("getRootNodes").treegrid("render")})},' + '\n' \
             'initNode:function(b){return this.each(function(){var c=a(this);c.treegrid("setTreeContainer",b.getTreeGridContainer.apply(this)),' + '\n' \
             'c.treegrid("getChildNodes").treegrid("initNode",b),' + '\n' \
             'c.treegrid("initExpander").treegrid("initIndent").treegrid("initEvents").treegrid("initState").treegrid("initChangeEvent").treegrid("initSettingsEvents")})},' + '\n' \
             'initChangeEvent:function(){var b=a(this);return b.on("change",function(){var b=a(this);b.treegrid("render"),b.treegrid("getSetting",' + '\n' \
             '"saveState")&&b.treegrid("saveState")}),b},initEvents:function(){var b=a(this);return b.on("collapse",function(){var b=a(this);' + '\n' \
             'b.removeClass("treegrid-expanded"),b.addClass("treegrid-collapsed")}),b.on("expand",function(){var b=a(this);b.removeClass("treegrid-collapsed"),' + '\n' \
             'b.addClass("treegrid-expanded")}),b},initSettingsEvents:function(){var b=a(this);return b.on("change",function(){var b=a(this);' + '\n' \
             '"function"==typeof b.treegrid("getSetting","onChange")&&b.treegrid("getSetting","onChange").apply(b)}),b.on("collapse",function(){var b=a(this);' + '\n' \
             '"function"==typeof b.treegrid("getSetting","onCollapse")&&b.treegrid("getSetting","onCollapse").apply(b)}),b.on("expand",function(){var b=a(this);' + '\n' \
             '"function"==typeof b.treegrid("getSetting","onExpand")&&b.treegrid("getSetting","onExpand").apply(b)}),b},initExpander:function()' + '\n' \
             '{var b=a(this),c=b.find("td").get(b.treegrid("getSetting","treeColumn")),d=b.treegrid("getSetting","expanderTemplate"),e=b.treegrid("getSetting",' + '\n' \
             '"getExpander").apply(this);return e&&e.remove(),a(d).prependTo(c).click(function(){a(a(this).closest("tr")).treegrid("toggle")}),b},' + '\n' \
             'initIndent:function(){var b=a(this);b.find(".treegrid-indent").remove();for(var c=b.treegrid("getSetting","indentTemplate"),' + '\n' \
             'd=b.find(".treegrid-expander"),e=b.treegrid("getDepth"),f=0;e>f;f++)a(c).insertBefore(d);return b},initState:function(){var b=a(this);' + '\n' \
             'return b.treegrid(b.treegrid("getSetting","saveState")&&!b.treegrid("isFirstInit")?"restoreState":"expanded"===b.treegrid("getSetting","initialState")?' + '\n' \
             '"expand":"collapse"),b},isFirstInit:function(){var b=a(this).treegrid("getTreeContainer");' + '\n' \
             'return void 0===b.data("first_init")&&b.data("first_init",void 0===a.cookie(b.treegrid("getSetting","saveStateName"))),' + '\n' \
             'b.data("first_init")},saveState:function(){var b=a(this);if("cookie"===b.treegrid("getSetting",' + '\n' \
             '"saveStateMethod")){var c=a.cookie(b.treegrid("getSetting","saveStateName"))||"",d=""===c?[]:c.split(","),' + '\n' \
             'e=b.treegrid("getNodeId");b.treegrid("isExpanded")?-1===a.inArray(e,d)&&d.push(e):b.treegrid("isCollapsed")&&-1!==a.inArray(e,d)&&d.splice(a.inArray(e,d),1),' + '\n' \
             'a.cookie(b.treegrid("getSetting","saveStateName"),d.join(","))}return b},restoreState:function(){var b=a(this);' + '\n' \
             'if("cookie"===b.treegrid("getSetting","saveStateMethod")){var c=a.cookie(b.treegrid("getSetting","saveStateName")).split(",");' + '\n' \
             'b.treegrid(-1!==a.inArray(b.treegrid("getNodeId"),c)?"expand":"collapse")}return b},getSetting:function(b)' + '\n' \
             '{return a(this).treegrid("getTreeContainer")?a(this).treegrid("getTreeContainer").data("settings")[b]:null},' + '\n' \
             'setSettings:function(b){a(this).treegrid("getTreeContainer").data("settings",b)},getTreeContainer:function(){return a(this).data("treegrid")},' + '\n' \
             'setTreeContainer:function(b){return a(this).data("treegrid",b)},getRootNodes:function(){return a(this).treegrid("getSetting",' + '\n' \
             '"getRootNodes").apply(this,[a(this).treegrid("getTreeContainer")])},getAllNodes:function(){return a(this).treegrid("getSetting",' + '\n' \
             '"getAllNodes").apply(this,[a(this).treegrid("getTreeContainer")])},isNode:function(){return null!==a(this).treegrid("getNodeId")},' + '\n' \
             'getNodeId:function(){return null===a(this).treegrid("getSetting","getNodeId")?null:a(this).treegrid("getSetting","getNodeId").apply(this)},' + '\n' \
             'getParentNodeId:function(){return a(this).treegrid("getSetting","getParentNodeId").apply(this)},getParentNode:function()' + '\n' \
             '{return null===a(this).treegrid("getParentNodeId")?null:a(this).treegrid("getSetting","getNodeById").apply(this,[a(this).treegrid("getParentNodeId"),' + '\n' \
             'a(this).treegrid("getTreeContainer")])},getChildNodes:function(){return a(this).treegrid("getSetting","getChildNodes").apply(this,' + '\n' \
             '[a(this).treegrid("getNodeId"),a(this).treegrid("getTreeContainer")])},getDepth:function(){return null===a(this).treegrid("getParentNode")?0:' + '\n' \
             'a(this).treegrid("getParentNode").treegrid("getDepth")+1},isRoot:function(){return 0===a(this).treegrid("getDepth")},isLeaf:function()' + '\n' \
             '{return 0===a(this).treegrid("getChildNodes").length},isLast:function(){if(a(this).treegrid("isNode")){var b=a(this).treegrid("getParentNode");' + '\n' \
             'if(null===b){if(a(this).treegrid("getNodeId")===a(this).treegrid("getRootNodes").last().treegrid("getNodeId"))return!0}' + '\n' \
             'else if(a(this).treegrid("getNodeId")===b.treegrid("getChildNodes").last().treegrid("getNodeId"))return!0}return!1},' + '\n' \
             'isFirst:function(){if(a(this).treegrid("isNode")){var b=a(this).treegrid("getParentNode");if(null===b)' + '\n' \
             '{if(a(this).treegrid("getNodeId")===a(this).treegrid("getRootNodes").first().treegrid("getNodeId"))return!0}' + '\n' \
             'else if(a(this).treegrid("getNodeId")===b.treegrid("getChildNodes").first().treegrid("getNodeId"))return!0}return!1},isExpanded:' + '\n' \
             'function(){return a(this).hasClass("treegrid-expanded")},isCollapsed:function(){return a(this).hasClass("treegrid-collapsed")},' + '\n' \
             'isOneOfParentsCollapsed:function(){var b=a(this);return b.treegrid("isRoot")?!1:b.treegrid("getParentNode").treegrid("isCollapsed")?!0:' + '\n' \
             'b.treegrid("getParentNode").treegrid("isOneOfParentsCollapsed")},expand:function(){return this.treegrid("isLeaf")||this.treegrid("isExpanded")?this:' + '\n' \
             '(this.trigger("expand"),this.trigger("change"),this)},expandAll:function(){var b=a(this);return b.treegrid("getRootNodes").treegrid("expandRecursive"),b},' + '\n' \
             'expandRecursive:function(){return a(this).each(function(){var b=a(this);b.treegrid("expand"),b.treegrid("isLeaf")||' + '\n' \
             'b.treegrid("getChildNodes").treegrid("expandRecursive")})},collapse:function(){return a(this).each(function(){var b=a(this);b.treegrid("isLeaf")||' + '\n' \
             'b.treegrid("isCollapsed")||(b.trigger("collapse"),b.trigger("change"))})},collapseAll:function(){var b=a(this);' + '\n' \
             'return b.treegrid("getRootNodes").treegrid("collapseRecursive"),b},collapseRecursive:function(){return a(this).each(function()' + '\n' \
             '{var b=a(this);b.treegrid("collapse"),b.treegrid("isLeaf")||b.treegrid("getChildNodes").treegrid("collapseRecursive")})},toggle:function()' + '\n' \
             '{var b=a(this);return b.treegrid(b.treegrid("isExpanded")?"collapse":"expand"),b},render:function(){return a(this).each(function(){var b=a(this);' + '\n' \
             'b.treegrid("isOneOfParentsCollapsed")?b.hide():b.show(),b.treegrid("isLeaf")||(b.treegrid("renderExpander"),' + '\n' \
             'b.treegrid("getChildNodes").treegrid("render"))})},renderExpander:function(){return a(this).each(function()' + '\n' \
             '{var b=a(this),c=b.treegrid("getSetting","getExpander").apply(this);c?b.treegrid("isCollapsed")?(c.removeClass(b.treegrid("getSetting",' + '\n' \
             '"expanderExpandedClass")),c.addClass(b.treegrid("getSetting","expanderCollapsedClass"))):(c.removeClass(b.treegrid("getSetting",' + '\n' \
             '"expanderCollapsedClass")),c.addClass(b.treegrid("getSetting","expanderExpandedClass"))):(b.treegrid("initExpander"),b.treegrid("renderExpander"))})}};' + '\n' \
             'a.fn.treegrid=function(c){return b[c]?b[c].apply(this,Array.prototype.slice.call(arguments,1)):' + '\n' \
             '"object"!=typeof c&&c?void a.error("Method with name "+c+" does not exists for jQuery.treegrid"):b.initTree.apply(this,arguments)},' + '\n' \
             'a.fn.treegrid.defaults={initialState:"expanded",saveState:!1,saveStateMethod:"cookie",saveStateName:"tree-grid-state",' + '\n' \
             'expanderTemplate:"<span class=''treegrid-expander''></span>",indentTemplate:"<span class=''treegrid-indent''></span>",' + '\n' \
             'expanderExpandedClass:"treegrid-expander-expanded",expanderCollapsedClass:"treegrid-expander-collapsed",' + '\n' \
             'treeColumn:0,getExpander:function(){return a(this).find(".treegrid-expander")},getNodeId:function(){var b=/treegrid-([A-Za-z0-9_-]+)/;' + '\n' \
             'return b.test(a(this).attr("class"))?b.exec(a(this).attr("class"))[1]:null},getParentNodeId:function(){var b=/treegrid-parent-([A-Za-z0-9_-]+)/;' + '\n' \
             'return b.test(a(this).attr("class"))?b.exec(a(this).attr("class"))[1]:null},getNodeById:function(a,b){var c="treegrid-"+a;return b.find("tr."+c)},' + '\n' \
             'getChildNodes:function(a,b){var c="treegrid-parent-"+a;return b.find("tr."+c)},getTreeGridContainer:function(){return a(this).closest("table")},' + '\n' \
             'getRootNodes:function(b){var c=a.grep(b.find("tr"),function(b){var c=a(b).attr("class"),d=/treegrid-([A-Za-z0-9_-]+)/,' + '\n' \
             'e=/treegrid-parent-([A-Za-z0-9_-]+)/;return d.test(c)&&!e.test(c)});return a(c)},getAllNodes:function(b)' + '\n' \
             '{var c=a.grep(b.find("tr"),function(b){var c=a(b).attr("class"),d=/treegrid-([A-Za-z0-9_-]+)/;return d.test(c)});' + '\n' \
             'return a(c)},onCollapse:null,onExpand:null,onChange:null}}(jQuery);' + '\n' \
         '</script>' + '\n' \
         '<script>' + '\n' \
             '$("#tree-1").treegrid({initialState: "collapsed"});' + '\n' \
         '</script>' + '\n' \
         '</body>' + '\n' \
         '</html>'
body = ''
length = pivot.shape[0]

level_memory = {} #словарь для позиций
keys = range(pos+1)
values = [0]*(pos+1)
for i in keys:
    level_memory[i] = values[i]

for i in range(length):
   grandparent = i + 1

   row_ = list(pivot.iloc[i, :]) #формирование строки таблицы
   for k, ind in enumerate(row_):
       if (ind != ' ') & (isinstance(ind, float)==False):
           level = k + 1
           break
   title = [x for x in row_ if (x != ' ') & (isinstance(ind, float)==False)]
   if len(title) == 1:
       row = [' ']*((n_levels+1)-len(title)) + row_[-len(month):]
   else:
       row = title[1:] + [' ']*((n_levels+1)-len(title)) + row_[-len(month):]
       title = [title[0]]
   if i + 1 != length: #проверка уровня следующей строки таблицы
        row_next_ = list(pivot.iloc[i+1, :])
        for k_next, ind_next in enumerate(row_next_):
           if (ind_next != ' ') & (isinstance(ind_next, float)==False):
               level_next = k_next + 1
               break
   else:
        level_next = level

   if level_next <= level: #упорядочивание элементов разворачивания
      expand = ''
      collapse = 'class="treegrid-expander">'
      parent = f' treegrid-parent-{previous}' #формирование зависимостей разворачивания
   else:
       level_count = title[0].count('__')
       expand = ' treegrid-collapsed'
       collapse = 'class="treegrid-expander treegrid-expander-collapsed">'

       if level_count == 0: #формирование зависимостей разворачивания
           parent = f' treegrid-parent-{grandparent}'
       else:
           parent = f' treegrid-parent-{level_memory[level_count-1]}'
       level_memory[level_count] = grandparent
       previous = grandparent

   if level == 1: #установка свойств разворачивания
       style=''
       parent = ''
   else:
       style = 'style="display: none;"'
   ids = grandparent - 1

   body += html_formatter.tr_expand(grandparent, parent, expand, #добавление строк в таблицу
                                    ids, style, collapse,
                                    title[0], n_items, row)

html_doc = header + body + footer #формирование полной таблицы

#Формирование html-таблицы
with open('./resources/table.html', 'w', newline='', encoding="utf-8") as file:
    file.write(html_doc)