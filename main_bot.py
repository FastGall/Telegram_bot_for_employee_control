# -*- coding: utf-8 -*-
import config
import telebot
import Generate_xls_file
from telebot import types
import xlrd, xlwt
from xlrd import open_workbook
from xlutils.copy import copy


bot = telebot.TeleBot('token')


@bot.message_handler(content_types=["text"])

def start(m):
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    keyboard.add(*[types.KeyboardButton(name) for name in ['Generate new table', 'Use table']])
    msg = bot.send_message(m.chat.id, 'Hello',
        reply_markup=keyboard)
    bot.register_next_step_handler(msg,name)

def autoriz(m):
    if(m.text == 'besT'):
        bot.send_message(m.chat.id, "Yep")
    else:
        bot.send_message(m.chat.id, "Неверный пароль")
        bot.register_next_step_handler(m,autoriz)


def name(m):
    if m.text == 'Generate new table':
        bot.send_message(m.chat.id, "print xls-file name for save")
        bot.register_next_step_handler(m,Generate_new_table)
    elif m.text == 'Use table':
        bot.send_message(m.chat.id, "[Day,Model,Time,Money,Employee]")
        bot.register_next_step_handler(m,Current)


def Generate_new_table(m):
    name_xls = str(m.text)  #Recieve to msg
    wb = xlwt.Workbook()
    ws = wb.add_sheet(m.text[0:9])
    ws.write(0, 0, 'Day')
    ws.write(0, 1, 'Model')
    ws.write(0, 2, 'Time')
    ws.write(0, 3, 'Money')
    ws.write(0, 4, 'Employee')
    wb.save(name_xls)
    bot.send_message(m.chat.id, "Create")

def Current(m):
    current_list = str(m.text) #Recieve msg
    current_list = current_list.split(',')
    #print(current_list,len(current_list))
    rb = open_workbook(current_list[0] + '.xls')
    wb = copy(rb)
    read = rb.sheet_by_index(0)
    sheet = wb.get_sheet(0) # Работа в текущем Exel-листе
    for i in range(len(current_list)):
        sheet.write(read.nrows, i, current_list[i])
       # print(current_list[i])
    wb.save(current_list[0] + '.xls')
    bot.send_message(m.chat.id, "Save")
    doc = open(current_list[0] + '.xls', 'rb')
    bot.send_document(m.chat.id, doc)
 


if __name__ == '__main__':
    bot.polling(none_stop=True)
