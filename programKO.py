import pandas as pd
import openpyxl
from add_size import ManSizes, WomanSizes
import os

FILENAME = 'data/Корпоративная одежда.xlsx'

TYPE_CLOTHES = {('поло',): 'Рубашка поло',
                ('руб', 'рубаш', 'рубашка'):
                  {('д/р', 'дл.рук',):
                       {('повседневаная', 'повс', 'повсед', 'голубая', 'голуб.', 'гол.'):
                            'Рубашка с длинным рукавом повседневная голубого цвета',
                       ('парад', 'пар.'):
                            'Рубашка с длинным рукавом парадная белого цвета'
                       },
                  ('к/р', 'кор.рук', 'кор.т.с.', 'кор.с.р', 'кор.р'):
                       {('повседневаная', 'повс', 'повсед', 'голубая', 'голуб.', 'гол.'):
                            'Рубашка с коротким рукавом повседневная голубого цвета',
                       ('парад', 'пар.'):
                            'Рубашка с коротким рукавом парадная белого цвета'
                       },
                   },
                  ('брюки',):
                      {('лет',):
                           'Брюки летние',
                       ('дем'):
                            'Брюки демисезонные',
                      },
                  ('Джемпер',): 'Джемпер трикотажный',
                  ('куртка',):
                    {
                        ('вет',): 'Куртка ветровка',
                        ('зим',): 'Куртка зимняя'
                    },
                  ('вет',): 'Куртка ветровка',
                ('Шапка', 'убор', 'голов', 'форменная'): 'головной убор зимний (трикотажная шапка)',
                    # {
                    #     ('трикотаж','зим',): 'головной убор зимний (трикотажная шапка)',
                    # },
                ('Жилет',):
                    {
                        ('утепленный', 'утепл',) : 'Жилет утепленный',
                    },
                ('юбка',):
                    {
                        ('демис',): 'Юбка демисезонная',
                        ('летняя',): 'Юбка летняя',
                    },
                ('платок',): 'Платок (шейный) или галстук-бант',
                ('кардиган',): 'Кардиган трикотажный',
              }


def search_clothes(cloth, iter_clothes=TYPE_CLOTHES):
    for type_cloth in iter_clothes:
        for el in type_cloth:
            if el.lower() in cloth.replace(' ', '').lower():
                iter_clothes = iter_clothes[type_cloth]
                if type(iter_clothes) is str:
                    return iter_clothes
                else:
                    return search_clothes(cloth, iter_clothes)

def clothes(tabs):
    '''
    Получение данных от бухгалтерии
    :param tabs:
    :return:
    '''
    df = pd.read_excel('data/бухКО.xlsx',
                       skiprows=9)
    tab_id = df['Unnamed: 2']
    skiped = df['Unnamed: 1']
    cloths = df['Unnamed: 11']
    dates = df['Unnamed: 10']
    for skip, tab_el, cloth, date in zip(skiped, tab_id, cloths, dates):
        # if tab_el == 19103726:
        #     if cloth == 'Рубашка муж. парадная к/р':
        #         print(cloth)
        if skip == '*' or tab_el not in tabs:
            continue
        else:
            re_cloth = search_clothes(cloth)
            if re_cloth:
                if tabs[tab_el]['clothes'].get(re_cloth):
                    tabs[tab_el]['clothes'][re_cloth]['date'].append(date.strftime('%d-%m-%Y'))
                else:
                    tabs[tab_el]['clothes'][re_cloth] = {
                        'date': [date.strftime('%d-%m-%Y')],
                    }
    return tabs


def set_data_woman(tabs):
    cells = {
        'Рубашка с длинным рукавом повседневная голубого цвета': (11, 12),
        'Рубашка с длинным рукавом парадная белого цвета': [13],
        'Брюки демисезонные': [14],
        'Юбка демисезонная': [15],
        'Платок (шейный) или галстук-бант': [16],
        'Кардиган трикотажный': [17],
        'Жилет утепленный': [18],
        'Куртка ветровка': (19,),
        'Куртка зимняя': (20,),
        'головной убор зимний (трикотажная шапка)': [21],
        'Рубашка с коротким рукавом повседневная голубого цвета': (22, 23),
        'Рубашка с коротким рукавом парадная белого цвета': [24],
        'Брюки летние': (25, ),
        'Юбка летняя': (26, ),
        'Рубашка поло': (27, )
    }
    wb_obj = openpyxl.load_workbook(FILENAME)
    sheet_obj = wb_obj['Женский комплект']
    set_data(cells=cells, sheet=sheet_obj, tabs=tabs)
    wb_obj.save(FILENAME)

def set_data(cells, sheet, tabs):
    '''
        Создание excel по данным от бугалтерии
        :param data:
        :return:
        '''

    max_cell = max(cells.values(), key=lambda v: v[0])[0]
    min_cell = min(cells.values(), key=lambda v: v[0])[0]
    tab_coll = 1
    surname_coll_num = 4
    recr_coll_num = 10
    ogre_coll = 2

    i = 6
    for tab in tabs:
        # if tab == 19009685:
        #     print()
        for j in (range(min_cell, max_cell + 1)):
            cell_obj = sheet.cell(row=i, column=j)
            cell_obj.value = 'не выдано'

        '''
        Задается инфа по человек
        '''
        tab_id = sheet.cell(row=i, column=tab_coll)
        tab_id.value = tab

        surname = sheet.cell(row=i, column=surname_coll_num)
        surname.value = tabs[tab]['surname']

        recr = sheet.cell(row=i, column=recr_coll_num)
        recr.value = tabs[tab]['recr_date']

        ogre = sheet.cell(row=i, column=ogre_coll)
        ogre.value = tabs[tab]['ogre']

        '''
        Задается инфа по одежде
        '''
        for cloth in tabs[tab]['clothes']:
            date = tabs[tab]['clothes'][cloth]['date']
            if cells.get(cloth):
                # if нужен если в буглатерии дадут одежду которую не должны были довать,
                # Например юбку для мужчины
                for num_coll, date_value in zip(cells[cloth], date):
                    cell_cloth = sheet.cell(row=i, column=num_coll)
                    cell_cloth.value = date_value
        i += 1

def set_data_mans(tabs):

    cells = {
        'Рубашка с длинным рукавом повседневная голубого цвета': (11, 12),
        'Рубашка с длинным рукавом парадная белого цвета': [13],
        'Брюки демисезонные': [14],
        'Джемпер трикотажный': [15],
        'Жилет утепленный': [16],
        'Куртка ветровка': [17],
        'Куртка зимняя': [18],
        'головной убор зимний (трикотажная шапка)': (19,),
        'Рубашка с коротким рукавом повседневная голубого цвета': (20, 21),
        'Рубашка с коротким рукавом парадная белого цвета': [22],
        'Брюки летние': (23,),
        'Рубашка поло': [24],
    }
    wb_obj = openpyxl.load_workbook('data/Шаблон Корпоративная одежда.xlsx')
    sheet_obj = wb_obj['Мужской комплект']
    set_data(cells=cells, sheet=sheet_obj, tabs=tabs)

    wb_obj.save(FILENAME)


def get_tab_id() -> dict:
    '''
    Получение всех нужных табельных номеров
    :return:
        list
    '''
    SETTINGS = {
        'type_people': 'рег.гор.пасс.марш.'
    }
    wb_obj = openpyxl.load_workbook('data/водители.xlsx')
    sheet_obj = wb_obj.active
    start = 2
    tab_coll = 1
    surname_coll = 2
    sex_coll = 5
    recr_date_coll = 6
    type_people_coll = 4
    orge_coll = 3

    tabs = {}
    for tab_num in range(start, sheet_obj.max_row + 1):
        type_people = sheet_obj.cell(row=tab_num, column=type_people_coll).value
        if SETTINGS['type_people'] in type_people:
            tabs[int(sheet_obj.cell(row=tab_num, column=tab_coll).value)] = {
                'sex': sheet_obj.cell(row=tab_num, column=sex_coll).value.upper(),
                'surname': sheet_obj.cell(row=tab_num, column=surname_coll).value,
                'recr_date': sheet_obj.cell(row=tab_num, column=recr_date_coll).value.strftime('%d-%m-%Y'),
                'ogre': sheet_obj.cell(row=tab_num, column=orge_coll).value,
                'clothes': {},
            }
        # if tabs.get(19009685):
        # if tabs.get(19009685):
        #     x = tabs[19009685]
        #     print()

    return tabs


def add_sizes():
    mans = ManSizes()
    womans = WomanSizes()

    mans.get_data(file='data/Корпоративная одежда старая.xlsx')
    mans.set_data()

    womans.get_data(file='data/Корпоративная одежда старая.xlsx')
    womans.set_data()


tabs = dict(sorted(get_tab_id().items()))
cloths_men = clothes({
    key: value for key, value in tabs.items() if value['sex'].upper() == 'мужской'.upper()
})
cloths_woman = clothes({
    key: value for key, value in tabs.items() if value['sex'].upper() == 'женский'.upper()
})

set_data_mans(cloths_men)
set_data_woman(cloths_woman)

add_sizes()


#TODO
'''
1. не корректное чтение файла бугалтерии
2. добавить один общий файл и перед зарузкой очищать его
'''
