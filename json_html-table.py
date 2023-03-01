#Библиотеки
import pandas as pd
import numpy as np
import warnings
import json
import os
warnings.filterwarnings('ignore')

#Функция изъятия первых n слов
def word_cut(df, col, n):
    new = df[col].str.split(' ')
    new = new.str[0:n]
    new = new.apply(lambda x: ' '.join(x))
    return new

#Методы для формирования html-таблицы
class html_formatter():
    def tr_expand(self, grandparent, parent, expand, ids, style, collapse, title, n_items, row):
        row_append = f'<tr class="treegrid-{grandparent}{parent}{expand}" id="{ids}" {style}>' \
                     f'<td>' \
                     f'<span {collapse} </span>' \
                     f'{title}' \
                     f'</td>'
        td = ''
        for n in range(n_items):
            td += f'<td>{row[n]}</td>'
        row_append += td + \
                      f'</tr>'
        return row_append

    def tr_build(self, all_cols):
        tr = '<tr id="tree-head-1">'
        for tag in all_cols:
            tr += '<th>' + tag + '</th>'
        tr += '</tr>'
        return tr
html_formatter = html_formatter()

#Открытие json
with open('./data.json') as fp:
    data = json.load(fp)

#1-ый уровень
fields = ['TITLE', 'ITEMS']
table = pd.json_normalize(data,max_level=1)[fields]
table = table.explode('ITEMS', ignore_index=True)
table = pd.concat([table, table['ITEMS'].apply(pd.Series)], axis=1)
fields = ['TITLE', 'title', 'ITEMS', 'STATUS_RESULT']
table = table[fields]
col_replace = {'TITLE': 'project', 'title': 'stage',
               'ITEMS': 'items', 'STATUS_RESULT': 'status_stage'}
table.rename(columns = col_replace, inplace = True)
table = table.iloc[:, np.r_[0,2,4,5]]
table = table.explode('items', ignore_index=True)
table = pd.concat([table, table['items'].apply(pd.Series)], axis=1)

#2-ой уровень
fields = ['project', 'stage', 'status_stage', 'title',
          'RESPONSOBLE_FIO', 'DURATION', 'BEGIN_DATE_FORMAT',
          'CLOSE_DATE_FORMAT', 'STATUS_RESULT', 'ITEMS', 'TASKS']
table = table[fields]
col_replace = {'title':'function', 'RESPONSOBLE_FIO': 'role_function',
               'DURATION': 'duration_function', 'BEGIN_DATE_FORMAT': 'start_function',
               'CLOSE_DATE_FORMAT': 'finish_function', 'STATUS_RESULT': 'status_function',
               'ITEMS': 'items', 'TASKS': 'tasks'}
table.rename(columns = col_replace, inplace = True)

#3-й уровень
fields = ['project', 'stage', 'function', 'status_stage', 'role_function',
         'duration_function', 'start_function', 'finish_function', 'tasks']
table = table[fields]
table = table.explode('tasks', ignore_index=True)
table = pd.concat([table, table['tasks'].apply(pd.Series)], axis=1)

#Удаление лишних столбцов
fields.append('TITLE')
table = table[fields]
col_replace = {'TITLE': 'task',
               'status_stage': 'status'}
table.rename(columns = col_replace, inplace = True)
table = table.drop('tasks', axis = 1)

#Создание структуры таблицы
table['decomposition_1'] = np.where(table['stage']==' ',
                                    'Задачи договора',
                                    'Этапы договора')
table['document'] = np.where(word_cut(table, 'function', 1).str[:-1].str.isnumeric()==True,
                             table['function'],
                             ' ')
table['function'] = np.where(table['document'] != ' ',
                             ' ',
                             table['function'])
table['decomposition_2'] = np.where(table['function']==' ',
                                    np.where(table['document']!=' ',
                                             'Документооборот',
                                             table['task']),
                                    'Функции')
table['task'] = np.where(table['task'] == table['decomposition_2'],
                         ' ',
                         table['task'])
table['decomposition_3'] = np.where(table['decomposition_2'] == 'Функции',
                                    table['function'],
                                    np.where(table['decomposition_2'] == 'Документооборот',
                                             table['document'],
                                             ' '))
table['stage'] = np.where(table['decomposition_1']=='Задачи договора',
                          table['decomposition_2'],
                          table['stage'])
table['decomposition_2'] = np.where(table['decomposition_2']==table['stage'],
                                    ' ',
                                    table['decomposition_2'])

#Добавление иерархии
col_order = ['project', 'decomposition_1', 'stage',
             'decomposition_2', 'decomposition_3', 'task',
             'role_function', 'duration_function', 'start_function',
             'finish_function']

pos = col_order.index('task') #вычисление числа уровней
table = table[col_order]
for i in range(pos):
    level = list(set(map(tuple,table.iloc[:,:(1+i)].values.tolist())))
    length = table.shape[1] - len(level[0])
    length = table.shape[1] - len(level[0])
    for j in level:
        row = list(j)
        table = table.append(pd.Series(row + [' ']*length,
                             index=list(table.columns)),
                             ignore_index=True)

#Составление сводной таблицы
pivot = table.pivot_table(index = col_order[:(pos+1)],
                          values = col_order[(pos+1):],
                          aggfunc ='sum')
pivot = pivot[~pivot.index.duplicated(keep='first')]
pivot = pivot[['duration_function', 'role_function',
               'start_function', 'finish_function']]
pivot['role_function'] = np.where((pivot['role_function'] == ' ') &
                                  (pivot['start_function'] != ' '),
                      '-',
                      pivot['role_function'])

#Сброс индекса
pivot.to_excel('./data.xlsx')
pivot = pd.read_excel('./data.xlsx').replace(np.nan, ' ')
os.remove('./data.xlsx')

#Добавление отступов
expands = list(pivot.columns)[:(pos+1)]
for k, col in enumerate(expands):
    pivot[col] = np.where(pivot[col] == ' ',
                                        ' ',
                                        '__'*k + pivot[col].astype(str))
#html-конвертация
n_items = pivot.shape[1] - pos - 1
header = '<!doctype html>' \
         '<html>' \
         '<head>' \
         '<meta charset="utf-8">' \
         '<title>Расписание</title>' \
         '<style>'\
              '.treegrid-indent {width: 16px;height: 16px;display: inline-block;' \
                                 'position: relative;}' \
              '.treegrid-expander {width: 16px;height: 16px;display: inline-block;' \
                                   'position: relative;}' \
              '.treegrid-expanded ' \
                '.treegrid-expander {background-color: green;cursor: pointer;}' \
              '.treegrid-collapsed ' \
                '.treegrid-expander {background-color: red;cursor: pointer;}'\
         '</style>'\
         '</head>' \
         '<body style="background-color: azure;">' \
         '<font face="Arial">'\
         '<table>'\
         '<tbody>'\
         '<tr>'\
         '<td>'\
         '<table id="tree-1" border="1">' \
         '<tbody>' \
         '<tr id="tree-head-1">' \
         '<th>Иерархия</th>' \
         '<th>Продолжительность</th>' \
         '<th>Ответственный</th>' \
         '<th>Начало</th>' \
         '<th>Окончание</th>' \
         '</tr>'
footer = '</tbody>'\
         '</table>' \
         '</font>'\
         '</td>'\
         '</tr>'\
         '</tbody>'\
         '</table>'\
         '<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.10.2/jquery.min.js"></script>'\
         '<script>'\
             '!function(a){var b={initTree:function(b){var c=a.extend({},this.treegrid.defaults,b);' \
             'return this.each(function(){var b=a(this);b.treegrid("setTreeContainer",a(this)),b.treegrid("setSettings",c),' \
             'c.getRootNodes.apply(this,[a(this)]).treegrid("initNode",c),b.treegrid("getRootNodes").treegrid("render")})},' \
             'initNode:function(b){return this.each(function(){var c=a(this);c.treegrid("setTreeContainer",b.getTreeGridContainer.apply(this)),' \
             'c.treegrid("getChildNodes").treegrid("initNode",b),' \
             'c.treegrid("initExpander").treegrid("initIndent").treegrid("initEvents").treegrid("initState").treegrid("initChangeEvent").treegrid("initSettingsEvents")})},' \
             'initChangeEvent:function(){var b=a(this);return b.on("change",function(){var b=a(this);b.treegrid("render"),b.treegrid("getSetting",' \
             '"saveState")&&b.treegrid("saveState")}),b},initEvents:function(){var b=a(this);return b.on("collapse",function(){var b=a(this);' \
             'b.removeClass("treegrid-expanded"),b.addClass("treegrid-collapsed")}),b.on("expand",function(){var b=a(this);b.removeClass("treegrid-collapsed"),' \
             'b.addClass("treegrid-expanded")}),b},initSettingsEvents:function(){var b=a(this);return b.on("change",function(){var b=a(this);' \
             '"function"==typeof b.treegrid("getSetting","onChange")&&b.treegrid("getSetting","onChange").apply(b)}),b.on("collapse",function(){var b=a(this);' \
             '"function"==typeof b.treegrid("getSetting","onCollapse")&&b.treegrid("getSetting","onCollapse").apply(b)}),b.on("expand",function(){var b=a(this);' \
             '"function"==typeof b.treegrid("getSetting","onExpand")&&b.treegrid("getSetting","onExpand").apply(b)}),b},initExpander:function()' \
             '{var b=a(this),c=b.find("td").get(b.treegrid("getSetting","treeColumn")),d=b.treegrid("getSetting","expanderTemplate"),e=b.treegrid("getSetting",' \
             '"getExpander").apply(this);return e&&e.remove(),a(d).prependTo(c).click(function(){a(a(this).closest("tr")).treegrid("toggle")}),b},' \
             'initIndent:function(){var b=a(this);b.find(".treegrid-indent").remove();for(var c=b.treegrid("getSetting","indentTemplate"),' \
             'd=b.find(".treegrid-expander"),e=b.treegrid("getDepth"),f=0;e>f;f++)a(c).insertBefore(d);return b},initState:function(){var b=a(this);' \
             'return b.treegrid(b.treegrid("getSetting","saveState")&&!b.treegrid("isFirstInit")?"restoreState":"expanded"===b.treegrid("getSetting","initialState")?' \
             '"expand":"collapse"),b},isFirstInit:function(){var b=a(this).treegrid("getTreeContainer");' \
             'return void 0===b.data("first_init")&&b.data("first_init",void 0===a.cookie(b.treegrid("getSetting","saveStateName"))),' \
             'b.data("first_init")},saveState:function(){var b=a(this);if("cookie"===b.treegrid("getSetting",' \
             '"saveStateMethod")){var c=a.cookie(b.treegrid("getSetting","saveStateName"))||"",d=""===c?[]:c.split(","),' \
             'e=b.treegrid("getNodeId");b.treegrid("isExpanded")?-1===a.inArray(e,d)&&d.push(e):b.treegrid("isCollapsed")&&-1!==a.inArray(e,d)&&d.splice(a.inArray(e,d),1),' \
             'a.cookie(b.treegrid("getSetting","saveStateName"),d.join(","))}return b},restoreState:function(){var b=a(this);' \
             'if("cookie"===b.treegrid("getSetting","saveStateMethod")){var c=a.cookie(b.treegrid("getSetting","saveStateName")).split(",");' \
             'b.treegrid(-1!==a.inArray(b.treegrid("getNodeId"),c)?"expand":"collapse")}return b},getSetting:function(b)' \
             '{return a(this).treegrid("getTreeContainer")?a(this).treegrid("getTreeContainer").data("settings")[b]:null},' \
             'setSettings:function(b){a(this).treegrid("getTreeContainer").data("settings",b)},getTreeContainer:function(){return a(this).data("treegrid")},' \
             'setTreeContainer:function(b){return a(this).data("treegrid",b)},getRootNodes:function(){return a(this).treegrid("getSetting",' \
             '"getRootNodes").apply(this,[a(this).treegrid("getTreeContainer")])},getAllNodes:function(){return a(this).treegrid("getSetting",' \
             '"getAllNodes").apply(this,[a(this).treegrid("getTreeContainer")])},isNode:function(){return null!==a(this).treegrid("getNodeId")},' \
             'getNodeId:function(){return null===a(this).treegrid("getSetting","getNodeId")?null:a(this).treegrid("getSetting","getNodeId").apply(this)},' \
             'getParentNodeId:function(){return a(this).treegrid("getSetting","getParentNodeId").apply(this)},getParentNode:function()' \
             '{return null===a(this).treegrid("getParentNodeId")?null:a(this).treegrid("getSetting","getNodeById").apply(this,[a(this).treegrid("getParentNodeId"),' \
             'a(this).treegrid("getTreeContainer")])},getChildNodes:function(){return a(this).treegrid("getSetting","getChildNodes").apply(this,' \
             '[a(this).treegrid("getNodeId"),a(this).treegrid("getTreeContainer")])},getDepth:function(){return null===a(this).treegrid("getParentNode")?0:' \
             'a(this).treegrid("getParentNode").treegrid("getDepth")+1},isRoot:function(){return 0===a(this).treegrid("getDepth")},isLeaf:function()' \
             '{return 0===a(this).treegrid("getChildNodes").length},isLast:function(){if(a(this).treegrid("isNode")){var b=a(this).treegrid("getParentNode");' \
             'if(null===b){if(a(this).treegrid("getNodeId")===a(this).treegrid("getRootNodes").last().treegrid("getNodeId"))return!0}' \
             'else if(a(this).treegrid("getNodeId")===b.treegrid("getChildNodes").last().treegrid("getNodeId"))return!0}return!1},' \
             'isFirst:function(){if(a(this).treegrid("isNode")){var b=a(this).treegrid("getParentNode");if(null===b)' \
             '{if(a(this).treegrid("getNodeId")===a(this).treegrid("getRootNodes").first().treegrid("getNodeId"))return!0}' \
             'else if(a(this).treegrid("getNodeId")===b.treegrid("getChildNodes").first().treegrid("getNodeId"))return!0}return!1},isExpanded:' \
             'function(){return a(this).hasClass("treegrid-expanded")},isCollapsed:function(){return a(this).hasClass("treegrid-collapsed")},' \
             'isOneOfParentsCollapsed:function(){var b=a(this);return b.treegrid("isRoot")?!1:b.treegrid("getParentNode").treegrid("isCollapsed")?!0:' \
             'b.treegrid("getParentNode").treegrid("isOneOfParentsCollapsed")},expand:function(){return this.treegrid("isLeaf")||this.treegrid("isExpanded")?this:' \
             '(this.trigger("expand"),this.trigger("change"),this)},expandAll:function(){var b=a(this);return b.treegrid("getRootNodes").treegrid("expandRecursive"),b},' \
             'expandRecursive:function(){return a(this).each(function(){var b=a(this);b.treegrid("expand"),b.treegrid("isLeaf")||' \
             'b.treegrid("getChildNodes").treegrid("expandRecursive")})},collapse:function(){return a(this).each(function(){var b=a(this);b.treegrid("isLeaf")||' \
             'b.treegrid("isCollapsed")||(b.trigger("collapse"),b.trigger("change"))})},collapseAll:function(){var b=a(this);' \
             'return b.treegrid("getRootNodes").treegrid("collapseRecursive"),b},collapseRecursive:function(){return a(this).each(function()' \
             '{var b=a(this);b.treegrid("collapse"),b.treegrid("isLeaf")||b.treegrid("getChildNodes").treegrid("collapseRecursive")})},toggle:function()' \
             '{var b=a(this);return b.treegrid(b.treegrid("isExpanded")?"collapse":"expand"),b},render:function(){return a(this).each(function(){var b=a(this);' \
             'b.treegrid("isOneOfParentsCollapsed")?b.hide():b.show(),b.treegrid("isLeaf")||(b.treegrid("renderExpander"),' \
             'b.treegrid("getChildNodes").treegrid("render"))})},renderExpander:function(){return a(this).each(function()' \
             '{var b=a(this),c=b.treegrid("getSetting","getExpander").apply(this);c?b.treegrid("isCollapsed")?(c.removeClass(b.treegrid("getSetting",' \
             '"expanderExpandedClass")),c.addClass(b.treegrid("getSetting","expanderCollapsedClass"))):(c.removeClass(b.treegrid("getSetting",' \
             '"expanderCollapsedClass")),c.addClass(b.treegrid("getSetting","expanderExpandedClass"))):(b.treegrid("initExpander"),b.treegrid("renderExpander"))})}};' \
             'a.fn.treegrid=function(c){return b[c]?b[c].apply(this,Array.prototype.slice.call(arguments,1)):' \
             '"object"!=typeof c&&c?void a.error("Method with name "+c+" does not exists for jQuery.treegrid"):b.initTree.apply(this,arguments)},' \
             'a.fn.treegrid.defaults={initialState:"expanded",saveState:!1,saveStateMethod:"cookie",saveStateName:"tree-grid-state",' \
             'expanderTemplate:"<span class=''treegrid-expander''></span>",indentTemplate:"<span class=''treegrid-indent''></span>",' \
             'expanderExpandedClass:"treegrid-expander-expanded",expanderCollapsedClass:"treegrid-expander-collapsed",' \
             'treeColumn:0,getExpander:function(){return a(this).find(".treegrid-expander")},getNodeId:function(){var b=/treegrid-([A-Za-z0-9_-]+)/;' \
             'return b.test(a(this).attr("class"))?b.exec(a(this).attr("class"))[1]:null},getParentNodeId:function(){var b=/treegrid-parent-([A-Za-z0-9_-]+)/;' \
             'return b.test(a(this).attr("class"))?b.exec(a(this).attr("class"))[1]:null},getNodeById:function(a,b){var c="treegrid-"+a;return b.find("tr."+c)},' \
             'getChildNodes:function(a,b){var c="treegrid-parent-"+a;return b.find("tr."+c)},getTreeGridContainer:function(){return a(this).closest("table")},' \
             'getRootNodes:function(b){var c=a.grep(b.find("tr"),function(b){var c=a(b).attr("class"),d=/treegrid-([A-Za-z0-9_-]+)/,' \
             'e=/treegrid-parent-([A-Za-z0-9_-]+)/;return d.test(c)&&!e.test(c)});return a(c)},getAllNodes:function(b)' \
             '{var c=a.grep(b.find("tr"),function(b){var c=a(b).attr("class"),d=/treegrid-([A-Za-z0-9_-]+)/;return d.test(c)});' \
             'return a(c)},onCollapse:null,onExpand:null,onChange:null}}(jQuery);'\
         '</script>'\
         '<script>'\
             '$("#tree-1").treegrid({initialState: "collapsed"});'\
         '</script>'\
         '</body>' \
         '</html>'
body = ''
length = pivot.shape[0]

level_memory = {} #создание словаря для позиций
keys = range(pos+1)
values = [0]*(pos+1)
for i in keys:
    level_memory[i] = values[i]
for i in range(length):
   grandparent = i + 1

   row_ = list(pivot.iloc[i, :]) #формирование строки таблицы
   for k, ind in enumerate(row_):
       if ind != ' ':
           level = k + 1
           break
   title = [x for x in row_ if x.isspace() == False]
   if (level != (pos+1)) & (len(title) == 1):
       row = [' ']*(pos-len(title))
   else:
       row = title[1:] + [' ']*(pos-len(title))
       title = [title[0]]

   if i + 1 != length: #проверка уровня следующей строки таблицы
        row_next_ = list(pivot.iloc[i+1, :])
        for k_next, ind_next in enumerate(row_next_):
           if ind_next != ' ':
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
with open('./table.html', 'w', newline='', encoding="utf-8") as file:
    file.write(html_doc)