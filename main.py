# -*- coding: utf-8 -*-
#####################################################################################################################
#                                                                                                                   #
# Python-скрипт для выгрузки данных по зарегистрированным пользователям из базы данных бота Comrade Major/SC Auth   #
#                                                                                                                   #
# MIT License                                                                                                       #
# Copyright (c) 2020 Michael Nikitenko                                                                              #
#                                                                                                                   #
#####################################################################################################################


import copy
import psycopg2
import xlwt
from configs import DB_CONFIG
from psycopg2.extras import DictCursor


def get_data_from_db():
    """
    Return list of users in dict format

    :return: list of dicts
    """
    conn = psycopg2.connect(dbname=DB_CONFIG['dbname'],
                            user=DB_CONFIG['user'],
                            password=DB_CONFIG['password'],
                            host=DB_CONFIG['host'])
    cursor = conn.cursor(cursor_factory=DictCursor)
    cursor.execute(f"SELECT * FROM {DB_CONFIG['table']}")
    res = cursor.fetchall()
    return res


def create_xls(data):
    """
    Recieve list of dicts with DB data and return create "Авторизация СЦ.xls" file.

    :param data: list of dicts with DB data
    """
    # Creating book and sheet
    book = xlwt.Workbook('utf8')
    sheet = book.add_sheet('Зарегистрированные пользователи')

    # Table header style
    header_style = xlwt.XFStyle()
    header_font = xlwt.Font()
    header_font.name = 'Arial'
    header_font.bold = True
    header_font.colour_index = xlwt.Style.colour_map['black']
    header_font.height = 260
    header_style.font = header_font

    # Data rows styles. Inherits Header style.
    data_style = xlwt.XFStyle()
    data_font = copy.deepcopy(header_font)
    data_font.bold = False
    data_font.height = 240
    data_style.font = data_font

    # TamTam id col data style. Inherits Data style
    tt_id_style = xlwt.XFStyle()
    tt_id_style.font = data_font
    tt_id_style.num_format_str = '0'

    # Creating borders
    borders = xlwt.Borders()
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1
    header_style.borders = data_style.borders = tt_id_style.borders = borders


    # Creating text alignment
    al = xlwt.Alignment()
    al.horz = al.HORZ_CENTER
    al.vert = al.VERT_CENTER
    header_style.alignment = al
    al.horz = al.HORZ_LEFT
    al.wrap = True
    data_style.alignment = al
    tt_id_style.alignment = al

    # Table Header declaration
    header = ['№', 'ТамТам id', 'ТамТам имя пользователя', 'ФИО', 'Департамент', 'Должность', 'Чаты']
    header_width = [1455, 4400, 8200, 10900, 14200, 24500, 3275]
    for i in range(len(header)):
        sheet.write(0, i, header[i], header_style)
        sheet.col(i).width = header_width[i]
    sheet.row(1).height = 2500

    # Adding data rows
    i = 1
    for d in data:
        sheet.row(i+1).height = 2500
        cols = [i, 'user_id', 'username', 'fio', 'dep', 'pos', 'chats']
        for col in range(len(cols)):
            if col == 1:
                style = tt_id_style
            else:
                style = data_style
            sheet.write(i, col, d[col], style)
        i = i + 1

    # Save XLS document
    sheet.portrait = False
    sheet.set_print_scaling(100)
    book.save('Авторизация СЦ.xls')


if __name__ == '__main__':
    create_xls(get_data_from_db())