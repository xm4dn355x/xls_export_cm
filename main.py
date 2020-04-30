# -*- coding: utf-8 -*-
#####################################################################################################################
#                                                                                                                   #
# Python-скрипт для выгрузки данных по зарегистрированным пользователям из базы данных бота Comrade Major/SC Auth   #
#                                                                                                                   #
# MIT License                                                                                                       #
# Copyright (c) 2020 Michael Nikitenko                                                                              #
#                                                                                                                   #
#####################################################################################################################

import psycopg2
import xlwt
from configs import DB_CONFIG
from psycopg2.extras import DictCursor


def get_data_from_db():
    conn = psycopg2.connect(dbname=DB_CONFIG['dbname'],
                            user=DB_CONFIG['user'],
                            password=DB_CONFIG['password'],
                            host=DB_CONFIG['host'])
    cursor = conn.cursor(cursor_factory=DictCursor)
    cursor.execute(f"SELECT * FROM {DB_CONFIG['table']}")
    res = cursor.fetchall()
    return res


def create_xls(data):
    print(data)
    book = xlwt.Workbook('utf8')
    sheet = book.add_sheet('Зарегистрированные пользователи')

    header_style = xlwt.XFStyle()
    header_font = xlwt.Font()
    header_font.name = 'Arial'
    header_font.bold = True
    header_font.colour_index = xlwt.Style.colour_map['black']
    header_font.height = 240
    header_style.font = header_font

    data_style = xlwt.XFStyle()
    data_font = xlwt.Font()
    data_font.name = 'Arial'
    data_font.bold = False
    data_font.colour_index = xlwt.Style.colour_map['black']
    data_font.height = 240
    data_style.font = data_font

    tt_id_style = xlwt.XFStyle()
    tt_id_style.font = data_font
    tt_id_style.num_format_str = '0'

    borders = xlwt.Borders()
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1
    header_style.borders = borders
    data_style.borders = borders
    tt_id_style.borders = borders

    al = xlwt.Alignment()
    al.horz = al.HORZ_CENTER
    al.vert = al.VERT_CENTER
    header_style.alignment = al
    al.horz = al.HORZ_LEFT
    al.wrap = True
    data_style.alignment = al
    tt_id_style.alignment = al


    # Объявляем заголовок
    sheet.write(0, 0, '№', header_style)
    sheet.write(0, 1, 'ТамТам id', header_style)
    sheet.write(0, 2, 'ТамТам имя пользователя', header_style)
    sheet.write(0, 3, 'ФИО', header_style)
    sheet.write(0, 4, 'Департамент', header_style)
    sheet.write(0, 5, 'Должность', header_style)
    sheet.write(0, 6, 'Чаты', header_style)
    sheet.row(1).height = 2500
    sheet.col(0).width = 1455
    sheet.col(1).width = 4400
    sheet.col(2).width = 8200
    sheet.col(3).width = 10900
    sheet.col(4).width = 14200
    sheet.col(5).width = 24500
    sheet.col(6).width = 3275

    # Добавляем данные из базы
    i = 1
    for d in data:
        sheet.row(i+1).height = 2500
        sheet.write(i, 0, i, data_style)
        sheet.write(i, 1, d['user_id'], tt_id_style)
        sheet.write(i, 2, d['username'], data_style)
        sheet.write(i, 3, d['fio'], data_style)
        sheet.write(i, 4, d['dep'], data_style)
        sheet.write(i, 5, d['pos'], data_style)
        sheet.write(i, 6, d['chats'], data_style)
        i = i + 1

    # Сохраняем документ
    sheet.portrait = False
    sheet.set_print_scaling(100)
    book.save('Авторизация СЦ.xls')


if __name__ == '__main__':
    data = get_data_from_db()
    xls = create_xls(data)