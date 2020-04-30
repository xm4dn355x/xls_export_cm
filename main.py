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
from configs import DB_CONFIG


def get_data_from_db():
    data = 'TEST'
    return data


def create_xls(data):
    print(data)


if __name__ == '__main__':
    data = get_data_from_db
    xls = create_xls(data)