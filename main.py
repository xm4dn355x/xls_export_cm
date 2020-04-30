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


# GLOBALS
HEADER = ['№', 'ТамТам id', 'ТамТам имя пользователя', 'ФИО', 'Департамент', 'Должность', 'Чаты']
COLS_WIDTH = [1455, 4400, 8200, 10900, 14200, 24500, 3275]
DBNAME, USER, PASSWORD, HOST, TABLE = DB_CONFIG['dbname'], DB_CONFIG['user'], DB_CONFIG['password'], DB_CONFIG['host'], \
                                      DB_CONFIG['table']


def get_data_from_db():
    """
    Return list of users in dict format

    :return: list of dicts
    """
    conn = psycopg2.connect(dbname=DBNAME, user=USER, password=PASSWORD, host=HOST)
    cursor = conn.cursor(cursor_factory=DictCursor)
    cursor.execute(f"SELECT * FROM {TABLE}")
    res = cursor.fetchall()
    return res


def create_xls(data):
    """
    Recieve list of dicts with DB data and create "Авторизация СЦ.xls" file using create_styles and render_table funcs.

    :param data: list of dicts with DB data
    """
    book = xlwt.Workbook('utf8')
    sheet = book.add_sheet('Зарегистрированные пользователи')
    styles = create_styles()
    header_style, data_style, tt_id_style = styles['header_style'], styles['data_style'], styles['tt_id_style']
    sheet, sheet.portrait = render_table(sheet, HEADER, COLS_WIDTH, data, header_style, data_style, tt_id_style), False
    sheet.set_print_scaling(100)
    book.save('Авторизация СЦ.xls')


def render_table(sheet, header, width, data, header_style, data_style, tt_id_style):
    """
    Rendering table from data and header parameters.

    :param sheet: obj sheet
    :param header: list of str header cols names
    :param width: list of int cols width
    :param data: list of dicts with DB data
    :param header_style: obj xlwt.XFStyle of table header
    :param data_style: obj xlwt.XFStyle of data rows
    :param tt_id_style: obj xlwt.XFStyle of TamTam request.user.user_id col
    :return: obj sheet
    """
    # Render table header
    for i in range(len(header)):
        sheet.write(0, i, header[i], header_style)
        sheet.col(i).width = width[i]
    sheet.row(1).height = 2500
    # Render table data
    i = 1
    for d in data:
        sheet.row(i + 1).height = 2500
        cols = [i, 'user_id', 'username', 'fio', 'dep', 'pos', 'chats']
        for col in range(len(cols)):
            if col == 0:
                sheet.write(i, col, i, tt_id_style)
            elif col == 1:
                sheet.write(i, col, d[col], tt_id_style)
            else:
                sheet.write(i, col, d[col], data_style)
        i = i + 1
    return sheet


def create_styles():
    """
    Creating styles for table cells

    :return: dict with xlwt.XFStyle objects
    """
    # Init all styles
    header_style, data_style, tt_id_style = xlwt.XFStyle(), xlwt.XFStyle(), xlwt.XFStyle()
    # Create fonts
    header_font = xlwt.Font()
    header_font.name, header_font.bold, header_font.colour_index = 'Arial', True, xlwt.Style.colour_map['black']
    header_font.height = 260
    data_font = copy.deepcopy(header_font)
    data_font.bold, data_font.height = False, 240
    # Set fonts to styles
    header_style.font = header_font
    data_style.font = tt_id_style.font = data_font
    # Set borders
    borders = xlwt.Borders()
    borders.left, borders.right, borders.top, borders.bottom = 1, 1, 1, 1
    header_style.borders = data_style.borders = tt_id_style.borders = borders
    # Set alignments
    al = xlwt.Alignment()
    al.horz, al.vert = al.HORZ_CENTER, al.VERT_CENTER
    header_style.alignment = al
    al.horz, al.wrap = al.HORZ_LEFT, True
    data_style.alignment = tt_id_style.alignment = al
    # Set integer cell format to tt_id_style
    tt_id_style.num_format_str = '0'
    return {'header_style': header_style, 'data_style': data_style, 'tt_id_style': tt_id_style}

if __name__ == '__main__':
    create_xls(get_data_from_db())
